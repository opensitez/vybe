// vybe-test: go/errors_join_unwrap/errors_is_nil_err
// origin: languages/go/tests/go/test_errors_join_unwrap.rs
// vybe-test-mode: compile

package main
import "errors"
var ErrT = errors.New("t")
func main() { _ = errors.Is(nil, ErrT) }
