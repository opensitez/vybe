// vybe-test: go/errors_package/errors_as_pointer_to_interface
// origin: languages/go/tests/go/test_errors_package.rs
// vybe-test-mode: compile

package main
import "errors"
type timeout interface { Timeout() bool }
func main() { var target timeout
_ = errors.As(errors.New("x"), &target) }
