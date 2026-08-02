// vybe-test: go/unicode_package/unicode_is_space_compile
// origin: languages/go/tests/go/test_unicode_package.rs
// vybe-test-mode: compile

package main
import "unicode"
func main() { _ = unicode.IsSpace('\t') }
