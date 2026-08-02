// vybe-test: go/strings_bytes_compare/strings_equal_fold_angstrom
// origin: languages/go/tests/go/test_strings_bytes_compare.rs
// vybe-test-mode: compile

package main
import "strings"
func main() { _ = strings.EqualFold("å", "Å") }
