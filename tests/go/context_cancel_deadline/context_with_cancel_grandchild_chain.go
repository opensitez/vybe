// vybe-test: go/context_cancel_deadline/context_with_cancel_grandchild_chain
// origin: languages/go/tests/go/test_context_cancel_deadline.rs
// vybe-test-mode: compile

package main
import "context"
func main() { a, ca := context.WithCancel(context.Background())
b, _ := context.WithCancel(a)
_, _ = context.WithCancel(b)
ca() }
