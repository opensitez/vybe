// vybe-test: go/cover_runtime_testing/trace_with_region
// origin: languages/go/tests/go/test_cover_runtime_testing.rs
// vybe-test-mode: compile

package main
import "runtime/trace"
import "context"
func main() { ctx, task := trace.NewTask(context.Background(), "job")
defer task.End()
_ = trace.WithRegion(ctx, "step", func() {}) }
