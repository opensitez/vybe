// vybe-test: go/context_cancel_deadline/context_with_cancel_done_open_before_cancel
// origin: languages/go/tests/go/test_context_cancel_deadline.rs

package main
import "fmt"
import "context"
func main() { ctx, _ := context.WithCancel(context.Background())
select { case <-ctx.Done(): fmt.Println("done")
default: fmt.Println("pending") } }
