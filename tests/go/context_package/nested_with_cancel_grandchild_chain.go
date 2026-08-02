// vybe-test: go/context_package/nested_with_cancel_grandchild_chain
// origin: languages/go/tests/go/test_context_package.rs
// vybe-test-mode: compile

package main
import "context"
func main() { a, ca := context.WithCancel(context.Background())
b, _ := context.WithCancel(a)
_, _ = context.WithCancel(b)
ca() }
