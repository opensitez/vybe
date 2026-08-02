// vybe-test: go/cover_runtime_testing/trace_start
// origin: languages/go/tests/go/test_cover_runtime_testing.rs
// vybe-test-mode: compile

package main
import "runtime/trace"
import "os"
func main() { _ = trace.Start(os.Stdout) }
