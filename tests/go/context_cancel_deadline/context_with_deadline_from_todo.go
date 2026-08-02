// vybe-test: go/context_cancel_deadline/context_with_deadline_from_todo
// origin: languages/go/tests/go/test_context_cancel_deadline.rs
// vybe-test-mode: compile

package main
import "context"
import "time"
func main() { _, cancel := context.WithDeadline(context.TODO(), time.Now().Add(time.Second))
defer cancel() }
