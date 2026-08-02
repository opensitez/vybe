// vybe-test: go/stdlib_net_mail_index/mail_parse_address
// origin: languages/go/tests/go/test_stdlib_net_mail_index.rs
// vybe-test-mode: compile

package main
import "net/mail"
func main() { _, _ = mail.ParseAddress("Go <go@example.com>") }
