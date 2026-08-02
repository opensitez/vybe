// vybe-test: go/context_cancel_deadline/context_with_deadline_child_inherits
// origin: languages/go/tests/go/test_context_cancel_deadline.rs
// vybe-test-mode: compile

package main
import "context"
import "time"
func main() { parent, pcancel := context.WithDeadline(context.Background(), time.Now().Add(time.Hour))
defer pcancel()
child, ccancel := context.WithCancel(parent)
defer ccancel()
_, ok := child.Deadline()
_ = ok }
