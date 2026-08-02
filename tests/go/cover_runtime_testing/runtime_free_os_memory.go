// vybe-test: go/cover_runtime_testing/runtime_free_os_memory
// origin: languages/go/tests/go/test_cover_runtime_testing.rs
// vybe-test-mode: compile

package main
import "runtime"
func main() { runtime.FreeOSMemory() }
