// vybe-test: go/strings_bytes_compare/strings_count_multibyte_substr
// origin: languages/go/tests/go/test_strings_bytes_compare.rs
// vybe-test-mode: compile

package main
import "strings"
func main() { _ = strings.Count("café", "é") }
