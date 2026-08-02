// vybe-test: go/context_cancel_deadline/context_deadline_returns_time_value
// origin: languages/go/tests/go/test_context_cancel_deadline.rs
// vybe-test-mode: compile

package main
import "context"
import "time"
func main() { ctx, cancel := context.WithDeadline(context.Background(), time.Now().Add(time.Second))
defer cancel()
t, ok := ctx.Deadline()
_ = t
_ = ok }
