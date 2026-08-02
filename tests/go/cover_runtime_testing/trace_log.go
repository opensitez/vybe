// vybe-test: go/cover_runtime_testing/trace_log
// origin: languages/go/tests/go/test_cover_runtime_testing.rs
// vybe-test-mode: compile

package main
import "runtime/trace"
import "context"
func main() { trace.Log(context.Background(), "key", "val") }
