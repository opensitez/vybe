// vybe-test: go/stdlib_net_mail_index/smtp_plain_auth
// origin: languages/go/tests/go/test_stdlib_net_mail_index.rs
// vybe-test-mode: compile

package main
import "net/smtp"
func main() { _ = smtp.PlainAuth("", "user", "pass", "localhost") }
