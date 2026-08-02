// vybe-test: go/context_package/with_deadline_absolute_time
// origin: languages/go/tests/go/test_context_package.rs
// vybe-test-mode: compile

package main
import "context"
import "time"
func main() { _, cancel := context.WithDeadline(context.Background(), time.Now().Add(time.Second))
defer cancel() }
