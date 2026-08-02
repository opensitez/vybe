// vybe-test: go/cover_net_extra/mail_address_string
// origin: languages/go/tests/go/test_cover_net_extra.rs
// vybe-test-mode: compile

package main
import "net/mail"
func main() { addr, _ := mail.ParseAddress("Go Team <go@example.com>")
_ = addr.String() }
