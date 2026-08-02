// vybe-test: go/unicode_utf16_norm/unicode_is_digit_devanagari
// origin: languages/go/tests/go/test_unicode_utf16_norm.rs
// vybe-test-mode: compile

package main
import "unicode"
func main() { _ = unicode.IsDigit('४') }
