// vybe-test: go/cover_net_extra/smtp_auth_cram_md5
// origin: languages/go/tests/go/test_cover_net_extra.rs
// vybe-test-mode: compile

package main
import "net/smtp"
func main() { c, _ := smtp.Dial("localhost:25")
if c != nil { defer c.Close()
_ = c.Auth(smtp.CRAMMD5Auth("user", "secret")) } }
