// vybe-test: go/errors_package/errors_is_on_custom_unwrap_chain
// origin: languages/go/tests/go/test_errors_package.rs
// vybe-test-mode: compile

package main
import "errors"
type link struct { next error }
func (l link) Error() string { return "link" }
func (l link) Unwrap() error { return l.next }
func main() { base := errors.New("base")
_ = errors.Is(link{next: base}, base) }
