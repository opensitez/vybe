// vybe-test: go/strings_ops_extended/split_after_empty_separator
// origin: languages/go/tests/go/test_strings_ops_extended.rs
// vybe-test-mode: compile

package main
import "strings"
func main() { _ = strings.SplitAfter("ab", "") }
