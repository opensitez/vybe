// vybe-test: go/strings_advanced/string_to_rune_slice
// origin: languages/go/tests/go/test_strings_advanced.rs
// vybe-test-mode: compile

package main
import "fmt"
func main() { s := "abc"
r := []rune(s)
_ = r }
