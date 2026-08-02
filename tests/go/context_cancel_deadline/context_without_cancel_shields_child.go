// vybe-test: go/context_cancel_deadline/context_without_cancel_shields_child
// origin: languages/go/tests/go/test_context_cancel_deadline.rs
// vybe-test-mode: compile

package main
import "context"
func main() { parent, pcancel := context.WithCancel(context.Background())
child := context.WithoutCancel(parent)
pcancel()
_ = child.Err() == nil }
