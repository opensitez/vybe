// vybe-test: go/context_cancel_deadline/context_with_value_chain_four_levels
// origin: languages/go/tests/go/test_context_cancel_deadline.rs
// vybe-test-mode: compile

package main
import "context"
func main() { c := context.Background()
c = context.WithValue(c, "l1", 1)
c = context.WithValue(c, "l2", 2)
c = context.WithValue(c, "l3", 3)
c = context.WithValue(c, "l4", 4)
_ = c.Value("l2") }
