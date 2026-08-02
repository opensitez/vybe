// vybe-test: go/errors_join_unwrap/errors_join_sentinel_var
// origin: languages/go/tests/go/test_errors_join_unwrap.rs
// vybe-test-mode: compile

package main
import "errors"
var ErrJ = errors.New("j")
func main() { _ = errors.Join(ErrJ, ErrJ) }
