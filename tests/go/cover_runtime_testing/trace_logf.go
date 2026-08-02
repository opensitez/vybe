// vybe-test: go/cover_runtime_testing/trace_logf
// origin: languages/go/tests/go/test_cover_runtime_testing.rs
// vybe-test-mode: compile

package main
import "runtime/trace"
import "context"
func main() { trace.Logf(context.Background(), "n=%d", 1) }
