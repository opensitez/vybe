// vybe-test: go/cover_net_extra/mail_message_body
// origin: languages/go/tests/go/test_cover_net_extra.rs
// vybe-test-mode: compile

package main
import "net/mail"
import "strings"
func main() { msg, _ := mail.ReadMessage(strings.NewReader("\r\nhello"))
_ = msg.Body }
