// vybe-test: go/unicode_utf16_norm/utf8_valid_string_empty
// origin: languages/go/tests/go/test_unicode_utf16_norm.rs
// vybe-test-mode: compile

package main
import "unicode/utf8"
func main() { _ = utf8.ValidString("") }
