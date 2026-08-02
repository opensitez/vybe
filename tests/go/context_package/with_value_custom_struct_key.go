// vybe-test: go/context_package/with_value_custom_struct_key
// origin: languages/go/tests/go/test_context_package.rs
// vybe-test-mode: compile

package main
import "context"
type ctxKey struct{}
func main() { _ = context.WithValue(context.Background(), ctxKey{}, 1) }
