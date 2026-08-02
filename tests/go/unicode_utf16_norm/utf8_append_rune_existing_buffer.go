// vybe-test: go/unicode_utf16_norm/utf8_append_rune_existing_buffer
// origin: languages/go/tests/go/test_unicode_utf16_norm.rs
// vybe-test-mode: compile

package main
import "unicode/utf8"
func main() { _ = utf8.AppendRune([]byte("a"), 'b') }
