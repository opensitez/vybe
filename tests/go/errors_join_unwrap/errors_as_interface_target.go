// vybe-test: go/errors_join_unwrap/errors_as_interface_target
// origin: languages/go/tests/go/test_errors_join_unwrap.rs
// vybe-test-mode: compile

package main
import "errors"
type timeout interface { Timeout() bool }
func main() { var t timeout
_ = errors.As(errors.New("x"), &t) }
