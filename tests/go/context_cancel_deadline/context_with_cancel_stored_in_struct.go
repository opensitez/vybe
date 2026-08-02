// vybe-test: go/context_cancel_deadline/context_with_cancel_stored_in_struct
// origin: languages/go/tests/go/test_context_cancel_deadline.rs
// vybe-test-mode: compile

package main
import "context"
type holder struct { ctx context.Context
cancel context.CancelFunc }
func main() { h := holder{}
h.ctx, h.cancel = context.WithCancel(context.Background())
defer h.cancel() }
