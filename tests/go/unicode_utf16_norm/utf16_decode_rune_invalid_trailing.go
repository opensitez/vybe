// vybe-test: go/unicode_utf16_norm/utf16_decode_rune_invalid_trailing
// origin: languages/go/tests/go/test_unicode_utf16_norm.rs
// vybe-test-mode: compile

package main
import "unicode/utf16"
func main() { _ = utf16.DecodeRune(0xDC00, 65535) }
