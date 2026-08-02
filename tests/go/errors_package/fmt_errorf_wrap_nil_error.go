// vybe-test: go/errors_package/fmt_errorf_wrap_nil_error
// origin: languages/go/tests/go/test_errors_package.rs
// vybe-test-mode: compile

package main
import "fmt"
func main() { _ = fmt.Errorf("wrap: %w", nil) }
