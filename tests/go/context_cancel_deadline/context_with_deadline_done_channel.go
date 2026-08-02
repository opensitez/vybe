// vybe-test: go/context_cancel_deadline/context_with_deadline_done_channel
// origin: languages/go/tests/go/test_context_cancel_deadline.rs
// vybe-test-mode: compile

package main
import "context"
import "time"
func main() { ctx, cancel := context.WithDeadline(context.Background(), time.Now().Add(time.Millisecond))
defer cancel()
<-ctx.Done() }
