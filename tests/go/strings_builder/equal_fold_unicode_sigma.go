// vybe-test: go/strings_builder/equal_fold_unicode_sigma
// origin: languages/go/tests/go/test_strings_builder.rs
// vybe-test-mode: compile

package main
import "strings"
func main() { _ = strings.EqualFold("σ", "Σ") }
