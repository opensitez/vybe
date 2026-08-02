// vybe-test: go/cover_net_extra/mail_read_message
// origin: languages/go/tests/go/test_cover_net_extra.rs
// vybe-test-mode: compile

package main
import "net/mail"
import "strings"
func main() { _, _ = mail.ReadMessage(strings.NewReader("Subject: hi\r\n\r\nbody")) }
