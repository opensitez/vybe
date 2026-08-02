// vybe-test: go/cover_runtime_testing/runtime_callers_frames
// origin: languages/go/tests/go/test_cover_runtime_testing.rs
// vybe-test-mode: compile

package main
import "runtime"
func main() { pcs := make([]uintptr, 4)
n := runtime.Callers(0, pcs)
frames := runtime.CallersFrames(pcs[:n])
_, _ = frames.Next() }
