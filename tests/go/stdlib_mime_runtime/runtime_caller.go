// vybe-test: go/stdlib_mime_runtime/runtime_caller
// origin: languages/go/tests/go/test_stdlib_mime_runtime.rs
// vybe-test-mode: compile

package main
import "runtime"
func main() { _, _, _, _ = runtime.Caller(0) }
