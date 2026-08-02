// vybe-test: go/fmt_sprintf_verbs/sprintf_go_string_verb
// origin: languages/go/tests/go/test_fmt_sprintf_verbs.rs
// vybe-test-mode: compile

package main
import "fmt"
func main() { fmt.Println(fmt.Sprintf("%#v", struct{ X int }{X: 1})) }
