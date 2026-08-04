// vybe-test: go/strings_advanced/rune_slice_to_string
// origin: languages/go/tests/go/test_strings_advanced.rs
// vybe-test-mode: compile

package main
import "fmt"
func main() { r := []rune{97, 98, 99}
s := string(r)
_ = s }
