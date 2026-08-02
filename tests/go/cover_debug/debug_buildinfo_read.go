// vybe-test: go/cover_debug/debug_buildinfo_read
// origin: languages/go/tests/go/test_cover_debug.rs
// vybe-test-mode: compile

package main
import "runtime/debug"
func main() { _, _ = debug.ReadBuildInfo() }
