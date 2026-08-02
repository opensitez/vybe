// vybe-test: go/cover_runtime_testing/runtime_caller
// origin: languages/go/tests/go/test_cover_runtime_testing.rs
// vybe-test-mode: compile

package main
import "runtime"
func main() { _, _, _, _ = runtime.Caller(0) }
