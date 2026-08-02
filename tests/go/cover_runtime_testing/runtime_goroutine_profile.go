// vybe-test: go/cover_runtime_testing/runtime_goroutine_profile
// origin: languages/go/tests/go/test_cover_runtime_testing.rs
// vybe-test-mode: compile

package main
import "runtime"
func main() { _, _ = runtime.GoroutineProfile(nil) }
