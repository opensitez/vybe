// vybe-test: go/errors_join_unwrap/errors_as_struct_pointer
// origin: languages/go/tests/go/test_errors_join_unwrap.rs
// vybe-test-mode: compile

package main
import "errors"
type E struct { N int }
func (e E) Error() string { return "e" }
func main() { var target E
_ = errors.As(E{N: 1}, &target) }
