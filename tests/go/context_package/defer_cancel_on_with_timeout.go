// vybe-test: go/context_package/defer_cancel_on_with_timeout
// origin: languages/go/tests/go/test_context_package.rs
// vybe-test-mode: compile

package main
import "context"
import "time"
func main() { ctx, cancel := context.WithTimeout(context.Background(), time.Hour)
defer cancel()
_ = ctx }
