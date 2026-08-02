// vybe-test: go/context_cancel_deadline/context_nested_with_value_and_cancel
// origin: languages/go/tests/go/test_context_cancel_deadline.rs
// vybe-test-mode: compile

package main
import "context"
func main() { base := context.WithValue(context.Background(), "id", "1")
mid, mcancel := context.WithCancel(base)
defer mcancel()
leaf := context.WithValue(mid, "step", 2)
_ = leaf.Value("id")
_ = leaf.Value("step") }
