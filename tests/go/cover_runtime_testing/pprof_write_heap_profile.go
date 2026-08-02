// vybe-test: go/cover_runtime_testing/pprof_write_heap_profile
// origin: languages/go/tests/go/test_cover_runtime_testing.rs
// vybe-test-mode: compile

package main
import "runtime/pprof"
import "os"
func main() { _ = pprof.WriteHeapProfile(os.Stdout) }
