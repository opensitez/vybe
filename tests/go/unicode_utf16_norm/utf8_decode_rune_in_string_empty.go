// vybe-test: go/unicode_utf16_norm/utf8_decode_rune_in_string_empty
// origin: languages/go/tests/go/test_unicode_utf16_norm.rs
// vybe-test-mode: compile

package main
import "unicode/utf8"
func main() { _, _ = utf8.DecodeRuneInString("") }
