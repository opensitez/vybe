// vybe-test: go/cover_runtime_testing/runtime_set_finalizer
// origin: languages/go/tests/go/test_cover_runtime_testing.rs
// vybe-test-mode: compile

package main
import "runtime"
type T struct{}
func main() { var t T
runtime.SetFinalizer(&t, func(*T) {}) }
