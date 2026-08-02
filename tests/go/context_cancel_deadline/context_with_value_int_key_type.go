// vybe-test: go/context_cancel_deadline/context_with_value_int_key_type
// origin: languages/go/tests/go/test_context_cancel_deadline.rs
// vybe-test-mode: compile

package main
import "context"
type ctxKey int
func main() { const k ctxKey = 0
_ = context.WithValue(context.Background(), k, "v") }
