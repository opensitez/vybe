// vybe-test: go/context_package/with_cancel_done_select_default_before_cancel
// origin: languages/go/tests/go/test_context_package.rs

package main
import "fmt"
import "context"
func main() { ctx, _ := context.WithCancel(context.Background())
select { case <-ctx.Done(): fmt.Println("closed")
default: fmt.Println("open") } }
