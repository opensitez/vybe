// vybe-test: go/errors_join_unwrap/errors_custom_is_on_join
// origin: languages/go/tests/go/test_errors_join_unwrap.rs
// vybe-test-mode: compile

package main
import "errors"
var ErrSpec = errors.New("spec")
type w struct { e error }
func (w w) Error() string { return w.e.Error() }
func (w w) Is(t error) bool { return t == ErrSpec }
func main() { _ = errors.Is(w{e: ErrSpec}, ErrSpec) }
