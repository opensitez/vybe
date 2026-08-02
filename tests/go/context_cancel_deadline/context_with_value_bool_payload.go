// vybe-test: go/context_cancel_deadline/context_with_value_bool_payload
// origin: languages/go/tests/go/test_context_cancel_deadline.rs
// vybe-test-mode: compile

package main
import "context"
func main() { _ = context.WithValue(context.Background(), "flag", true).Value("flag") }
