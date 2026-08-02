// vybe-test: go/context_cancel_deadline/context_with_cancel_done_closed_after_cancel
// origin: languages/go/tests/go/test_context_cancel_deadline.rs

package main
import "fmt"
import "context"
func main() { ctx, cancel := context.WithCancel(context.Background())
cancel()
select { case <-ctx.Done(): fmt.Println("done")
default: fmt.Println("pending") } }
