// vybe-test: go/fmt_sprintf_verbs/sprintf_slice_brackets
// origin: languages/go/tests/go/test_fmt_sprintf_verbs.rs
// vybe-test-mode: compile

package main
import "fmt"
func main() { fmt.Println(fmt.Sprintf("%v", []int{1,2})) }
