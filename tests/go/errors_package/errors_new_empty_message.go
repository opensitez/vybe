// vybe-test: go/errors_package/errors_new_empty_message
// origin: languages/go/tests/go/test_errors_package.rs
// vybe-test-mode: compile

package main
import "errors"
func main() { _ = errors.New("") }
