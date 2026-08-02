// vybe-test: go/errors_join_unwrap/errors_unwrap_custom
// origin: languages/go/tests/go/test_errors_join_unwrap.rs
// vybe-test-mode: compile

package main
import "errors"
type chain struct { next error }
func (c chain) Error() string { return "chain" }
func (c chain) Unwrap() error { return c.next }
func main() { _ = errors.Unwrap(chain{next: errors.New("inner")}) }
