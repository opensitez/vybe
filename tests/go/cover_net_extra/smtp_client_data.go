// vybe-test: go/cover_net_extra/smtp_client_data
// origin: languages/go/tests/go/test_cover_net_extra.rs
// vybe-test-mode: compile

package main
import "net/smtp"
func main() { c, _ := smtp.Dial("localhost:25")
if c != nil { defer c.Close()
_, _ = c.Data() } }
