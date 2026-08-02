// vybe-test: go/strings_bytes_compare/strings_compare_long_unicode
// origin: languages/go/tests/go/test_strings_bytes_compare.rs
// vybe-test-mode: compile

package main
import "strings"
func main() { _ = strings.Compare("α", "β") }
