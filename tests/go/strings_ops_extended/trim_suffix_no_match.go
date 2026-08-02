// vybe-test: go/strings_ops_extended/trim_suffix_no_match
// origin: languages/go/tests/go/test_strings_ops_extended.rs
// vybe-test-mode: compile

package main
import "strings"
func main() { _, _ = strings.TrimSuffix("go", "lang") }
