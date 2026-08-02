// vybe-test: go/errors_join_unwrap/errors_join_with_fmt_wrap
// origin: languages/go/tests/go/test_errors_join_unwrap.rs
// vybe-test-mode: compile

package main
import "fmt"
import "errors"
func main() { _ = errors.Join(fmt.Errorf("w: %w", errors.New("inner"))) }
