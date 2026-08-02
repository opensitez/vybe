// vybe-test: go/fmt_errors_print/fscanf_scan_three_fields
// origin: languages/go/tests/go/test_fmt_errors_print.rs
// vybe-test-mode: compile

package main
import "fmt"
import "strings"
func main() { var a, b, c int
_, _ = fmt.Fscan(strings.NewReader("1 2 3"), &a, &b, &c) }
