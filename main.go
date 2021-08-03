package main

import (
	"bytes"
	"crypto/tls"
	"encoding/json"
	"errors"
	"flag"
	"fmt"
	gotempalte "html/template"
	"io"
	"log"
	"net/mail"
	"os"
	"strings"
	"time"

	"github.com/tealeg/xlsx"
	"gopkg.in/gomail.v2"
)

type ContentProvider func(data interface{}) (string, func(writer io.Writer) error)

type Send struct {
	SendTo string
	Subject string
	Content *string
	Meta map[string]string
}

type Config struct {
	Host string `json:"host"`
	Port int `json:"port"`
	Username string `json:"username"`
	Password string `json:"password"`
	From string `json:"from"`
	Interval int64 `json:"interval"`
	Sender string `json:"sender"`
}

var (
	config string

	content string
	template string

	debug bool
	help bool
)

func logDebug(format string, v ...interface{}) {
	if debug {
		log.Printf(fmt.Sprintf("[DEBUG] %s", format), v...)
	}
}

func init() {
	flag.StringVar(&config, "config", "config.json", "配置文件")

	flag.StringVar(&content, "content", "", "邮件内容")
	flag.StringVar(&template, "template", "", "邮件模板")

	flag.BoolVar(&debug, "debug", false, "debug mode print detail log")
	flag.BoolVar(&help, "help", false, "print help info")
}

func main() {
	logDebug("参数列表: %s", os.Args[1:])

	flag.Parse()

	if help {
		usage()
		return
	}

	if flag.NArg() < 1 {
		log.Fatal("请提供 Excel 数据文件")
	}

	if len(config) == 0 {
		log.Fatal("请指定配置文件")
	}

	data, err := readFileContent(config)
	if err != nil {
		log.Fatalf("读取配置文件失败：%s", err)
	}
	var cfg Config
	if err = json.Unmarshal(data, &cfg); err != nil {
		log.Fatalf("读取配置文件失败：%s", err)
	}

	logDebug("解析完配置内容：%+v", &cfg)

	contentProvider, err := getContentProvider(content, template)
	if err != nil {
		log.Fatal(err)
	}

	file := flag.Arg(0)

	list, err := loadSendList(file)
	if err != nil {
		log.Fatalf("处理 Excel 文件失败：%s", err)
	}

	logDebug("处理完成，有 %d 条待发送邮件", len(list))

	sendEmails(&cfg, list, contentProvider)
}

func sendEmails(cfg *Config, list []*Send, contentProvider ContentProvider) {

	sender, err := getSender(cfg)
	if err != nil {
		log.Fatalf("创建 Sender 失败：%s", err)
	}

	defer func() {
		if closer, ok := sender.(io.Closer); ok {
			closer.Close()
		}
	}()

	m := gomail.NewMessage()

	for _, s := range list {
		m.SetHeader("From", cfg.From)
		m.SetHeader("To", s.SendTo)
		m.SetHeader("Subject", s.Subject)

		if s.Content != nil {
			m.SetBody(detectContentType([]byte(*s.Content)), *s.Content)
		} else {
			ct, content := contentProvider(s.Meta)
			m.AddAlternativeWriter(ct, content)
		}

		if err := gomail.Send(sender, m); err != nil {
			log.Printf("发送失败 %s -> %s: %v", s.SendTo, s.Content, err)
		}
		logDebug("To: %s, 发送成功", s.SendTo)
		m.Reset()

		if cfg.Interval > 0 {
			time.Sleep(time.Millisecond * time.Duration(cfg.Interval))
		}
	}
}

func loadSendList(file string) ([]*Send, error) {
	excel, err := xlsx.OpenFile(file)
	if err != nil {
		return nil, err
	}

	if len(excel.Sheets) == 0 || len(excel.Sheets[0].Rows) == 0 {
		return nil, errors.New("空表格")
	}

	rows := excel.Sheets[0].Rows

	maybeHeader := rows[0]
	skipHeader, rowParser, err := getRowParser(maybeHeader)

	if err != nil {
		return nil, err
	}
	if skipHeader {
		rows = rows[1:]
	}

	list := []*Send{}

	for i, row := range rows {
		send, err := rowParser(row)
		if err != nil {
			return nil, errors.New(fmt.Sprintf("解析第 %d 行出错，%s", i + 1, err))
		}
		list = append(list, send)
	}

	return list, nil
}

func getRowParser(first *xlsx.Row) (bool, func(row *xlsx.Row) (*Send, error), error) {
	if len(first.Cells) < 2 {
		return false, nil, errors.New("最少需要两列(SendTo, Subject)")
	}

	headerRow := false

	for _, cell := range first.Cells {
		if strings.Contains("SendTo, Subject, Content", cell.Value) {
			headerRow = true
			break
		}
	}

	if headerRow {
		logDebug("Header Excel")

		handlers := map[int]func(val string, send *Send) error {}

		for i, cell := range first.Cells {
			switch cell.Value {
			case "SendTo":
				handlers[i] = func(val string, send *Send) error {
					if !validEmailAddress(val) {
						return errors.New(fmt.Sprintf("无效的收件人: %s", val))
					}
					send.SendTo = val
					return nil
				}
			case "Subject":
				handlers[i] = func(val string, send *Send) error {
					if len(val) == 0 {
						return errors.New("标题不能为空")
					}
					send.Subject = val
					return nil
				}
			case "Content":
				handlers[i] = func(val string, send *Send) error {
					if len(val) != 0 {
						send.Content = &val
					}
					return nil
				}
			default:
				logDebug("Meta Cell: %s", cell.Value)
				key := cell.Value
				handlers[i] = func(val string, send *Send) error {
					if len(val) > 0 {
						if send.Meta == nil {
							send.Meta = map[string]string{}
						}
						send.Meta[key] = val
					}
					return nil
				}
			}
		}

		return true, func(row *xlsx.Row) (*Send, error) {
			var send Send
			for i, cell := range row.Cells {
				if handler, ok := handlers[i]; ok {
					if err := handler(cell.Value, &send); err != nil {
						return nil, err
					}
				} else {
					return nil, errors.New("数据与表头对不上")
				}
			}
			return &send, nil
		}, nil

	} else {
		return false, func(row *xlsx.Row) (*Send, error) {

			if len(row.Cells) < 2 {
				return nil, errors.New("最少需要两列(SendTo, Subject)")
			}
			sendTo := row.Cells[0].Value
			if len(sendTo) == 0 || !validEmailAddress(sendTo) {
				return nil, errors.New(fmt.Sprintf("无效的收件人: %s", sendTo))
			}
			subject := row.Cells[1].Value
			if len(subject) == 0 {
				return nil, errors.New("邮件标题不能为空")
			}

			var content *string

			if len(row.Cells) > 2 {
				content = &row.Cells[2].Value
			}
			return &Send{SendTo: sendTo, Subject: subject, Content: content}, nil
		}, nil
	}
}

func getContentProvider(content, template string) (ContentProvider, error) {

	if len(content) == 0 && len(template) == 0 {
		return nil, errors.New("邮件内容或邮件模板必须指定一个")
	} else if len(content) != 0 && len(template) != 0 {
		return nil, errors.New("邮件内容或邮件模板只能指定一个")
	}

	if len(content) > 0 {
		logDebug("从 %s 中读取邮件内容", content)
		data, err := readFileContent(content)
		if err != nil {
			log.Fatalf("读取邮件内容文件失败：%s", err)
		}
		contentType := detectContentType(data)

		logDebug("使用邮件内容 %s: %s", contentType, string(data))

		return func(_data interface{}) (s string, f func(writer io.Writer) error) {
			return contentType, func(w io.Writer) error {
				_, err := io.WriteString(w, string(data))
				return err
			}
		}, nil

	} else {
		logDebug("从 %s 中读取邮件内容", template)
		data, err := readFileContent(template)
		if err != nil {
			log.Fatalf("读取邮件模板文件失败：%s", err)
		}
		t, err := gotempalte.New("email").Parse(string(data))
		if err != nil {
			log.Fatalf("解析邮件模板失败：%s", err)
		}
		contentType := detectContentType(data)

		logDebug("使用邮件模板 %s: %s", contentType, string(data))

		return func(data interface{}) (s string, f func(writer io.Writer) error) {
			logDebug("Template Data: %+v", data)
			return contentType, func(w io.Writer) error {
				return t.Execute(w, data)
			}
		}, nil
	}
}

func getSender(cfg *Config) (gomail.Sender, error) {
	switch cfg.Sender {
	case "fake":
		return gomail.SendFunc(func(from string, to []string, msg io.WriterTo) error {
			var buffer bytes.Buffer
			if _, err := msg.WriteTo(&buffer); err != nil {
				return err
			}
			log.Printf("%s Send email to %s: %s", from, to, buffer.String())
			return nil
		}), nil
	default:
		d := gomail.NewDialer(cfg.Host, cfg.Port, cfg.Username, cfg.Password)
		d.TLSConfig = &tls.Config{InsecureSkipVerify: true}
		return d.Dial()
	}
}

func readFileContent(filename string) (data []byte, err error) {
	file, err := os.Open(filename)
	if err != nil {
		return nil, err
	}
	return io.ReadAll(file)
}

func detectContentType(data []byte) string {
	idx1 := bytes.IndexByte(data, '<')
	idx2 := bytes.IndexByte(data, '>')

	if idx1 > -1 && idx2 > -1 {
		return "text/html"
	} else {
		return "text/plain"
	}
}

func validEmailAddress(addr string) bool {
	a, err := mail.ParseAddress(addr)
	return err == nil && a != nil
}

func usage() {
	fmt.Print(`
	批量邮件发送助手 v0.1

	使用方式：
		email-sender.exe [--debug] --config config.json [--content content.txt | --template template.tpl] test.xlsx

	选项说明：
	
	--debug 打印详细信息
	
	--help 显示此帮助信息

	--config 指定配置文件路径

	--content 指定邮件内容文件路径，文件内容可以包含 html； 与 --template 选项冲突，只能使用一个
	
	--template 指定邮件内容模板文件路径，文件内容可以包含 html； 与 --content 选项冲突，只能使用一个

	配置文件参考：
	{
	  "host": "smtp.163.com",
	  "port": 465,
	  "username": "helloworld_hyx@163.com",
	  "password": "--PASSWORLD--",
	  "from": "helloworld_hyx@163.com",
	  "interval": 200,
	  "sender": "fake"
	}
	
	邮件内容文件：
	
	邮件模板文件：
	模板文件中可以使用 {{ .Xxxx }} 的语法访问 Excel 文件中自定义的其他列

	Excel 源文件说明：
	目前支持两种格式
	固定格式：
	SendTo, Subject, Content
	+-------------------------------------------------------+
	| helloworld_hyx@163.com  | Subject1 | Optional content |
	+-------------------------------------------------------+
	| helloworld_hyx@qq.com  | Subject2                     |
	+-------------------------------------------------------+
	
	* Content 是可以选的，如果内容不为空则替代 --content / --template 选项指定的内容

	或者带表头的格式：
	+---------------+----------+---------+-----+
	|    SendTo     | Subject  | Content | Xxx |
	+---------------+----------+---------+-----+
	| abc@hello.com | Subject1 |         |   1 |
	| def@hello.com | Subject2 | abc     |   2 |
	+---------------+----------+---------+-----+

	* 表格头（SendTo，Subject，Content）为内置名称，除了 Content 外，都必须提供，顺序无所谓
	* Content 是可以选的，如果内容不为空则替代 --content / --template 选项指定的内容
	* Xxx 可以是任意的，并且可以有多个，可以在模板文件中访问
`)
}
