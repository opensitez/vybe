// vybe-test: go/interface_assertion_extended/assert_error_interface_from_concrete_compile
// origin: languages/go/tests/go/test_interface_assertion_extended.rs
// vybe-test-mode: compile

package main
import "errors"
func main() { err := errors.New("x")
_, ok := err.(error)
_ = ok }
