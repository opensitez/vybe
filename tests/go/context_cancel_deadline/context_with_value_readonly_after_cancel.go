// vybe-test: go/context_cancel_deadline/context_with_value_readonly_after_cancel
// origin: languages/go/tests/go/test_context_cancel_deadline.rs
// vybe-test-mode: compile

package main
import "context"
func main() { ctx, cancel := context.WithCancel(context.WithValue(context.Background(), "k", 9))
cancel()
_ = ctx.Value("k") }
