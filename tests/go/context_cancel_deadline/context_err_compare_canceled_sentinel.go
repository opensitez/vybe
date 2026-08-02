// vybe-test: go/context_cancel_deadline/context_err_compare_canceled_sentinel
// origin: languages/go/tests/go/test_context_cancel_deadline.rs
// vybe-test-mode: compile

package main
import "context"
func main() { ctx, cancel := context.WithCancel(context.Background())
cancel()
_ = ctx.Err() == context.Canceled }
