// vybe-test: go/context_cancel_deadline/context_background_value_always_nil
// origin: languages/go/tests/go/test_context_cancel_deadline.rs
// vybe-test-mode: compile

package main
import "context"
func main() { _ = context.Background().Value("any") == nil }
