// vybe-test: go/context_cancel_deadline/context_err_nil_on_fresh_with_cancel
// origin: languages/go/tests/go/test_context_cancel_deadline.rs
// vybe-test-mode: compile

package main
import "context"
func main() { ctx, cancel := context.WithCancel(context.Background())
defer cancel()
_ = ctx.Err() == nil }
