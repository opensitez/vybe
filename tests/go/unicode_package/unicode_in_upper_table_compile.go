// vybe-test: go/unicode_package/unicode_in_upper_table_compile
// origin: languages/go/tests/go/test_unicode_package.rs
// vybe-test-mode: compile

package main
import "unicode"
func main() { _ = unicode.In('A', unicode.Upper) }
