// vybe-test: go/cover_runtime_testing/debug_set_max_threads
// origin: languages/go/tests/go/test_cover_runtime_testing.rs
// vybe-test-mode: compile

package main
import "runtime/debug"
func main() { _ = debug.SetMaxThreads(10000) }
