// vybe-test: go/stdlib_net_mail_index/smtp_send_mail
// origin: languages/go/tests/go/test_stdlib_net_mail_index.rs
// vybe-test-mode: compile

package main
import "net/smtp"
func main() { _ = smtp.SendMail("localhost:25", nil, "from@example.com", []string{"to@example.com"}, []byte("body")) }
