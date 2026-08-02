// vybe-test: go/cover_net_extra/mail_message_write_to
// origin: languages/go/tests/go/test_cover_net_extra.rs
// vybe-test-mode: compile

package main
import "bytes"
import "net/mail"
import "strings"
func main() { msg, _ := mail.ReadMessage(strings.NewReader("Subject: x\r\n\r\n"))
var buf bytes.Buffer
_, _ = msg.WriteTo(&buf) }
