// vybe-test: go/errors_join_unwrap/errorf_wrap_nil
// origin: languages/go/tests/go/test_errors_join_unwrap.rs
// vybe-test-mode: compile

package main
import "fmt"
func main() { _ = fmt.Errorf("wrap: %w", nil) }
