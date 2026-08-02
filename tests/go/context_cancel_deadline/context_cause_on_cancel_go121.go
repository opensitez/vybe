// vybe-test: go/context_cancel_deadline/context_cause_on_cancel_go121
// origin: languages/go/tests/go/test_context_cancel_deadline.rs
// vybe-test-mode: compile

package main
import "context"
import "errors"
func main() { ctx, cancel := context.WithCancelCause(context.Background())
cancel(errors.New("stop"))
_ = context.Cause(ctx) }
