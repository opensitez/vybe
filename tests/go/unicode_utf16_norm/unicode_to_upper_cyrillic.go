// vybe-test: go/unicode_utf16_norm/unicode_to_upper_cyrillic
// origin: languages/go/tests/go/test_unicode_utf16_norm.rs
// vybe-test-mode: compile

package main
import "unicode"
func main() { _ = unicode.ToUpper('ж') }
