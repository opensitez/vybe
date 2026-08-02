// vybe-test: go/cover_net_extra/mail_word_encoder
// origin: languages/go/tests/go/test_cover_net_extra.rs
// vybe-test-mode: compile

package main
import "net/mail"
type Encoder = mail.WordEncoder
func main() { var enc Encoder
_ = enc.Encode("hello world") }
