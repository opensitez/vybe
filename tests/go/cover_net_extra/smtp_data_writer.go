// vybe-test: go/cover_net_extra/smtp_data_writer
// origin: languages/go/tests/go/test_cover_net_extra.rs
// vybe-test-mode: compile

package main
import "net/smtp"
func main() { c, _ := smtp.Dial("localhost:25")
if c != nil { defer c.Close()
w, _ := c.Data()
if w != nil { _, _ = w.Write([]byte("hello")) } } }
