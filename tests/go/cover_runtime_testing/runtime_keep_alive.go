// vybe-test: go/cover_runtime_testing/runtime_keep_alive
// origin: languages/go/tests/go/test_cover_runtime_testing.rs
// vybe-test-mode: compile

package main
import "runtime"
func main() { x := 1
runtime.KeepAlive(x) }
