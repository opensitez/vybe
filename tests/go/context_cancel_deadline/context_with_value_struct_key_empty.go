// vybe-test: go/context_cancel_deadline/context_with_value_struct_key_empty
// origin: languages/go/tests/go/test_context_cancel_deadline.rs
// vybe-test-mode: compile

package main
import "context"
type key struct{}
func main() { _ = context.WithValue(context.Background(), key{}, true) }
