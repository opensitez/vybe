// vybe-test: go/unicode_utf16_norm/utf8_rune_len_out_of_range
// origin: languages/go/tests/go/test_unicode_utf16_norm.rs
// vybe-test-mode: compile

package main
import "unicode/utf8"
func main() { _ = utf8.RuneLen(0x110000) }
