// vybe-test: go/errors_package/errors_join_three_constituents
// origin: languages/go/tests/go/test_errors_package.rs
// vybe-test-mode: compile

package main
import "errors"
func main() { _ = errors.Join(errors.New("a"), errors.New("b"), errors.New("c")) }
