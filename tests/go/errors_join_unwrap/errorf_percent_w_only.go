// vybe-test: go/errors_join_unwrap/errorf_percent_w_only
// origin: languages/go/tests/go/test_errors_join_unwrap.rs
// vybe-test-mode: compile

package main
import "fmt"
import "errors"
func main() { _ = fmt.Errorf("%w", errors.New("cause")) }
