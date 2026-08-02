// vybe-test: go/context_cancel_deadline/context_done_channel_receive_compile
// origin: languages/go/tests/go/test_context_cancel_deadline.rs
// vybe-test-mode: compile

package main
import "context"
func main() { ctx, cancel := context.WithCancel(context.Background())
go func() { cancel() }()
_ = <-ctx.Done() }
