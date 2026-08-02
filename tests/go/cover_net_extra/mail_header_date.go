// vybe-test: go/cover_net_extra/mail_header_date
// origin: languages/go/tests/go/test_cover_net_extra.rs
// vybe-test-mode: compile

package main
import "net/mail"
import "strings"
func main() { msg, _ := mail.ReadMessage(strings.NewReader("Date: Mon, 02 Jan 2006 15:04:05 MST\r\n\r\n"))
_, _ = msg.Header.Date() }
