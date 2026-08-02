// vybe-test: go/cover_runtime_testing/debug_write_heap_dump
// origin: languages/go/tests/go/test_cover_runtime_testing.rs
// vybe-test-mode: compile

package main
import "runtime/debug"
import "os"
func main() { debug.WriteHeapDump(os.Stdout.Fd()) }
