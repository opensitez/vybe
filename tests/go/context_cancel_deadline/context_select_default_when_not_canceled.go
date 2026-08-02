// vybe-test: go/context_cancel_deadline/context_select_default_when_not_canceled
// origin: languages/go/tests/go/test_context_cancel_deadline.rs
// vybe-test-mode: compile

package main
import "context"
func main() { ctx, cancel := context.WithCancel(context.Background())
defer cancel()
select { case <-ctx.Done(): default: _ = 1 } }
