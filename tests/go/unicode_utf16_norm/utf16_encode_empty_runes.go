// vybe-test: go/unicode_utf16_norm/utf16_encode_empty_runes
// origin: languages/go/tests/go/test_unicode_utf16_norm.rs
// vybe-test-mode: compile

package main
import "unicode/utf16"
func main() { _ = utf16.Encode([]rune{}) }
