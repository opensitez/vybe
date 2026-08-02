// vybe-test: go/context_package/background_done_nil_falls_through_select_default
// origin: languages/go/tests/go/test_context_package.rs

package main
import "fmt"
import "context"
func main() { ctx := context.Background()
select { case <-ctx.Done(): fmt.Println("closed")
default: fmt.Println("open") } }
