// vybe-test: go/errors_join_unwrap/errors_join_many
// origin: languages/go/tests/go/test_errors_join_unwrap.rs
// vybe-test-mode: compile

package main
import "errors"
func main() { _ = errors.Join(errors.New("a"), errors.New("b"), errors.New("c"), errors.New("d"), errors.New("e")) }
