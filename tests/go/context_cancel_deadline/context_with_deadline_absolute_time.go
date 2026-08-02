// vybe-test: go/context_cancel_deadline/context_with_deadline_absolute_time
// origin: languages/go/tests/go/test_context_cancel_deadline.rs
// vybe-test-mode: compile

package main
import "context"
import "time"
func main() { _, cancel := context.WithDeadline(context.Background(), time.Now().Add(10*time.Second))
defer cancel() }
