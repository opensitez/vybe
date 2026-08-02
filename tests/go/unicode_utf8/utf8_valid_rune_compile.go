// vybe-test: go/unicode_utf8/utf8_valid_rune_compile
// origin: languages/go/tests/go/test_unicode_utf8.rs
// vybe-test-mode: compile

package main
import "unicode/utf8"
func main() { _ = utf8.ValidRune('🙂') }
