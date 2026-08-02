// vybe-test: go/errors_join_unwrap/errors_unwrap_multi_no_unwrap
// origin: languages/go/tests/go/test_errors_join_unwrap.rs
// vybe-test-mode: compile

package main
import "errors"
type plain struct{}
func (plain) Error() string { return "p" }
func main() { _ = errors.Unwrap(plain{}) }
