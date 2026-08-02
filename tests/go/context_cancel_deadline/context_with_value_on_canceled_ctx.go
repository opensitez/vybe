// vybe-test: go/context_cancel_deadline/context_with_value_on_canceled_ctx
// origin: languages/go/tests/go/test_context_cancel_deadline.rs
// vybe-test-mode: compile

package main
import "context"
func main() { parent, cancel := context.WithCancel(context.Background())
cancel()
_ = context.WithValue(parent, "k", 1) }
