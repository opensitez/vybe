// vybe-test: go/cover_runtime_testing/runtime_callers
// origin: languages/go/tests/go/test_cover_runtime_testing.rs
// vybe-test-mode: compile

package main
import "runtime"
func main() { pcs := make([]uintptr, 8)
_, _ = runtime.Callers(0, pcs) }
