// vybe-test: go/unicode_utf8/utf8_append_rune_compile
// origin: languages/go/tests/go/test_unicode_utf8.rs
// vybe-test-mode: compile

package main
import "unicode/utf8"
func main() { buf := []byte{}
_ = utf8.AppendRune(buf, 'a') }
