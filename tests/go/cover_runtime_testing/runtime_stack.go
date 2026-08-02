// vybe-test: go/cover_runtime_testing/runtime_stack
// origin: languages/go/tests/go/test_cover_runtime_testing.rs
// vybe-test-mode: compile

package main
import "runtime"
func main() { buf := make([]byte, 64)
_ = runtime.Stack(buf, false) }
