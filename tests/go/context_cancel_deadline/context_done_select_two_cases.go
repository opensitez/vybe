// vybe-test: go/context_cancel_deadline/context_done_select_two_cases
// origin: languages/go/tests/go/test_context_cancel_deadline.rs
// vybe-test-mode: compile

package main
import "context"
func main() { ctx, cancel := context.WithCancel(context.Background())
defer cancel()
ch := make(chan int)
select { case <-ctx.Done(): case <-ch: } }
