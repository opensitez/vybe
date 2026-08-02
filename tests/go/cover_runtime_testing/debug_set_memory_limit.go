// vybe-test: go/cover_runtime_testing/debug_set_memory_limit
// origin: languages/go/tests/go/test_cover_runtime_testing.rs
// vybe-test-mode: compile

package main
import "runtime/debug"
func main() { _ = debug.SetMemoryLimit(-1) }
