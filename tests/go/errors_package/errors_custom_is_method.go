// vybe-test: go/errors_package/errors_custom_is_method
// origin: languages/go/tests/go/test_errors_package.rs
// vybe-test-mode: compile

package main
import "errors"
var ErrSpecial = errors.New("special")
type wrapper struct { err error }
func (w wrapper) Error() string { return w.err.Error() }
func (w wrapper) Unwrap() error { return w.err }
func (w wrapper) Is(target error) bool { return target == ErrSpecial }
func main() { _ = errors.Is(wrapper{err: ErrSpecial}, ErrSpecial) }
