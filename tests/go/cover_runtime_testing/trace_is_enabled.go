// vybe-test: go/cover_runtime_testing/trace_is_enabled
// origin: languages/go/tests/go/test_cover_runtime_testing.rs
// vybe-test-mode: compile

package main
import "runtime/trace"
func main() { _ = trace.IsEnabled() }
