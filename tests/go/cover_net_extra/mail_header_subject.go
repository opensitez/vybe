// vybe-test: go/cover_net_extra/mail_header_subject
// origin: languages/go/tests/go/test_cover_net_extra.rs
// vybe-test-mode: compile

package main
import "net/mail"
import "strings"
func main() { msg, _ := mail.ReadMessage(strings.NewReader("Subject: test\r\n\r\n"))
_ = msg.Header.Subject() }
