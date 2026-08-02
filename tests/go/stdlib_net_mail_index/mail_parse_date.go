// vybe-test: go/stdlib_net_mail_index/mail_parse_date
// origin: languages/go/tests/go/test_stdlib_net_mail_index.rs
// vybe-test-mode: compile

package main
import "net/mail"
func main() { _, _ = mail.ParseDate("Mon, 02 Jan 2006 15:04:05 MST") }
