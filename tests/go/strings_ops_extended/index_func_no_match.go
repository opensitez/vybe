// vybe-test: go/strings_ops_extended/index_func_no_match
// origin: languages/go/tests/go/test_strings_ops_extended.rs
// vybe-test-mode: compile

package main
import "strings"
func main() { _ = strings.IndexFunc("abc", func(r rune) bool { return r == 'z' }) }
