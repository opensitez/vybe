// vybe-test: go/fmt_sprintf_verbs/sprintf_error_verb
// origin: languages/go/tests/go/test_fmt_sprintf_verbs.rs
// vybe-test-mode: compile

package main
import "fmt"
import "errors"
func main() { fmt.Println(fmt.Sprintf("%v", errors.New("boom"))) }
