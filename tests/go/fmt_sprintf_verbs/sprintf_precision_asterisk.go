// vybe-test: go/fmt_sprintf_verbs/sprintf_precision_asterisk
// origin: languages/go/tests/go/test_fmt_sprintf_verbs.rs
// vybe-test-mode: compile

package main
import "fmt"
func main() { fmt.Println(fmt.Sprintf("%.*f", 1, 2.25)) }
