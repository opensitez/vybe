// vybe-test: go/cover_net_extra/mail_header_add
// origin: languages/go/tests/go/test_cover_net_extra.rs
// vybe-test-mode: compile

package main
import "net/mail"
import "strings"
func main() { msg, _ := mail.ReadMessage(strings.NewReader("\r\n"))
msg.Header.Add("Received", "from localhost") }
