// vybe-test: go/context_package/select_ctx_done_or_work_channel
// origin: languages/go/tests/go/test_context_package.rs
// vybe-test-mode: compile

package main
import "context"
func main() { ctx, cancel := context.WithCancel(context.Background())
defer cancel()
work := make(chan int, 1)
select { case <-ctx.Done(): case work <- 1: } }
