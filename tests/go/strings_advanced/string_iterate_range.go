// vybe-test: go/strings_advanced/string_iterate_range
// origin: languages/go/tests/go/test_strings_advanced.rs
// vybe-test-mode: compile

package main
import "fmt"
func main() { s := "abc"
for i, r := range s { _ = i
_ = r
} }
