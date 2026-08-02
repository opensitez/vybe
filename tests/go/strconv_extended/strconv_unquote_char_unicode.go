// vybe-test: go/strconv_extended/strconv_unquote_char_unicode
// origin: languages/go/tests/go/test_strconv_extended.rs
// vybe-test-mode: compile

package main
import "strconv"
func main() { _, _, _ = strconv.UnquoteChar(`\u03BB`, `\`) }
