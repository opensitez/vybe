// vybe-test: go/unicode_utf16_norm/utf16_encode_rune_bmp_boundary
// origin: languages/go/tests/go/test_unicode_utf16_norm.rs
// vybe-test-mode: compile

package main
import "unicode/utf16"
func main() { _, _ = utf16.EncodeRune(0xFFFF) }
