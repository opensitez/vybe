// vybe-test: go/cover_net_extra/mail_parse_address_list
// origin: languages/go/tests/go/test_cover_net_extra.rs
// vybe-test-mode: compile

package main
import "net/mail"
func main() { _, _ = mail.ParseAddressList("Alice <a@example.com>, Bob <b@example.com>") }
