// vybe-test: go/context_cancel_deadline/context_with_timeout_from_todo
// origin: languages/go/tests/go/test_context_cancel_deadline.rs
// vybe-test-mode: compile

package main
import "context"
import "time"
func main() { _, cancel := context.WithTimeout(context.TODO(), time.Second)
defer cancel() }
