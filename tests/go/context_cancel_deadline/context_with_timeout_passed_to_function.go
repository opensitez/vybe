// vybe-test: go/context_cancel_deadline/context_with_timeout_passed_to_function
// origin: languages/go/tests/go/test_context_cancel_deadline.rs
// vybe-test-mode: compile

package main
import "context"
import "time"
func work(ctx context.Context) { _ = ctx.Err() }
func main() { ctx, cancel := context.WithTimeout(context.Background(), time.Second)
defer cancel()
work(ctx) }
