// vybe-test: go/context_cancel_deadline/context_with_timeout_nested_child
// origin: languages/go/tests/go/test_context_cancel_deadline.rs
// vybe-test-mode: compile

package main
import "context"
import "time"
func main() { parent, pcancel := context.WithTimeout(context.Background(), time.Minute)
defer pcancel()
_, ccancel := context.WithTimeout(parent, time.Second)
defer ccancel() }
