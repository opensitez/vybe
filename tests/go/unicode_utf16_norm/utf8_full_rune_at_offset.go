// vybe-test: go/unicode_utf16_norm/utf8_full_rune_at_offset
// origin: languages/go/tests/go/test_unicode_utf16_norm.rs
// vybe-test-mode: compile

package main
import "unicode/utf8"
func main() { _ = utf8.FullRuneAt([]byte("世"), 0) }
