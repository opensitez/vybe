// vybe-test: go/cover_net_extra/smtp_dial
// origin: languages/go/tests/go/test_cover_net_extra.rs
// vybe-test-mode: compile

package main
import "net/smtp"
func main() { _, _ = smtp.Dial("localhost:25") }
