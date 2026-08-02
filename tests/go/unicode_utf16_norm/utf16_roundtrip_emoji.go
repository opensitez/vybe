// vybe-test: go/unicode_utf16_norm/utf16_roundtrip_emoji
// origin: languages/go/tests/go/test_unicode_utf16_norm.rs
// vybe-test-mode: compile

package main
import "unicode/utf16"
func main() { u := utf16.Encode([]rune("🎉"))
_ = utf16.Decode(u) }
