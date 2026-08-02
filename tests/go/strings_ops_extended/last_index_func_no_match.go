// vybe-test: go/strings_ops_extended/last_index_func_no_match
// origin: languages/go/tests/go/test_strings_ops_extended.rs
// vybe-test-mode: compile

package main
import "strings"
func main() { _ = strings.LastIndexFunc("abc", func(r rune) bool { return r == 'z' }) }
