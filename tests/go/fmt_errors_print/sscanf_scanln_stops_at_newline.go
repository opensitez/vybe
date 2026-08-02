// vybe-test: go/fmt_errors_print/sscanf_scanln_stops_at_newline
// origin: languages/go/tests/go/test_fmt_errors_print.rs
// vybe-test-mode: compile

package main
import "fmt"
func main() { var s string
_, _ = fmt.Sscanln("line\nrest", &s) }
