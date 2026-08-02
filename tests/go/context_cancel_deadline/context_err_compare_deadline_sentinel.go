// vybe-test: go/context_cancel_deadline/context_err_compare_deadline_sentinel
// origin: languages/go/tests/go/test_context_cancel_deadline.rs
// vybe-test-mode: compile

package main
import "context"
import "time"
func main() { ctx, cancel := context.WithTimeout(context.Background(), 0)
defer cancel()
_ = ctx.Err() == context.DeadlineExceeded }
