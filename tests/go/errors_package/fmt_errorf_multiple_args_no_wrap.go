// vybe-test: go/errors_package/fmt_errorf_multiple_args_no_wrap
// origin: languages/go/tests/go/test_errors_package.rs
// vybe-test-mode: compile

package main
import "fmt"
func main() { _ = fmt.Errorf("%s %d", "err", 1) }
