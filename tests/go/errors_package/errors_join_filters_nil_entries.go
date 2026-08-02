// vybe-test: go/errors_package/errors_join_filters_nil_entries
// origin: languages/go/tests/go/test_errors_package.rs
// vybe-test-mode: compile

package main
import "errors"
func main() { _ = errors.Join(nil, errors.New("only"), nil) }
