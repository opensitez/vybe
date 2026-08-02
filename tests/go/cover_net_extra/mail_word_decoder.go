// vybe-test: go/cover_net_extra/mail_word_decoder
// origin: languages/go/tests/go/test_cover_net_extra.rs
// vybe-test-mode: compile

package main
import "net/mail"
type Decoder = mail.WordDecoder
func main() { var dec Decoder
_, _ = dec.Decode("hello") }
