// vybe-test: go/cover_net_extra/smtp_cram_md5_auth
// origin: languages/go/tests/go/test_cover_net_extra.rs
// vybe-test-mode: compile

package main
import "net/smtp"
func main() { _ = smtp.CRAMMD5Auth("user", "secret") }
