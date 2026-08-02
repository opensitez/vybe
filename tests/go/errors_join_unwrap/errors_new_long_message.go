// vybe-test: go/errors_join_unwrap/errors_new_long_message
// origin: languages/go/tests/go/test_errors_join_unwrap.rs
// vybe-test-mode: compile

package main
import "errors"
func main() { _ = errors.New("a long error message for compile smoke") }
