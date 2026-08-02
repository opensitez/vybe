// vybe-test: go/unicode_utf16_norm/utf8_rune_count_in_string_emoji
// origin: languages/go/tests/go/test_unicode_utf16_norm.rs
// vybe-test-mode: compile

package main
import "unicode/utf8"
func main() { _ = utf8.RuneCountInString("🙂") }
