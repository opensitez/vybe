// vybe-test: go/fmt_errors_print/fscanf_scanln_from_reader
// origin: languages/go/tests/go/test_fmt_errors_print.rs
// vybe-test-mode: compile

package main
import "fmt"
import "strings"
func main() { var s string
_, _ = fmt.Fscanln(strings.NewReader("one two"), &s) }
