// vybe-test: go/stdlib_mime_runtime/debug_read_build_info
// origin: languages/go/tests/go/test_stdlib_mime_runtime.rs
// vybe-test-mode: compile

package main
import "runtime/debug"
func main() { _, _ = debug.ReadBuildInfo() }
