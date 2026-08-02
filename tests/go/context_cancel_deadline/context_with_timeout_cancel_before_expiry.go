// vybe-test: go/context_cancel_deadline/context_with_timeout_cancel_before_expiry
// origin: languages/go/tests/go/test_context_cancel_deadline.rs
// vybe-test-mode: compile

package main
import "context"
import "time"
func main() { ctx, cancel := context.WithTimeout(context.Background(), time.Hour)
cancel()
_ = ctx.Err() == context.Canceled }
