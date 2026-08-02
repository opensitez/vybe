// vybe-test: go/cover_runtime_testing/debug_read_build_info
// origin: languages/go/tests/go/test_cover_runtime_testing.rs
// vybe-test-mode: compile

package main
import "runtime/debug"
func main() { _, _ = debug.ReadBuildInfo() }
