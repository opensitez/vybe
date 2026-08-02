// vybe-test: go/strings_ops_extended/strings_clone_independent
// origin: languages/go/tests/go/test_strings_ops_extended.rs
// vybe-test-mode: compile

package main
import "strings"
func main() { s := "go"
c := strings.Clone(s)
_ = c }
