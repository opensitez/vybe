// vybe-test: go/fmt_errors_print/errorf_octal_and_width
// origin: languages/go/tests/go/test_fmt_errors_print.rs
// vybe-test-mode: compile

package main
import "fmt"
func main() { _ = fmt.Errorf("#%04o", 7) }
