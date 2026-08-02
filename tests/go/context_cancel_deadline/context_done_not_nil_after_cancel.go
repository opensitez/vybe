// vybe-test: go/context_cancel_deadline/context_done_not_nil_after_cancel
// origin: languages/go/tests/go/test_context_cancel_deadline.rs
// vybe-test-mode: compile

package main
import "context"
func main() { ctx, cancel := context.WithCancel(context.Background())
cancel()
_ = ctx.Done() != nil }
