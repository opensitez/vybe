// vybe-test: go/context_cancel_deadline/context_with_timeout_one_second
// origin: languages/go/tests/go/test_context_cancel_deadline.rs
// vybe-test-mode: compile

package main
import "context"
import "time"
func main() { _, cancel := context.WithTimeout(context.Background(), time.Second)
defer cancel() }
