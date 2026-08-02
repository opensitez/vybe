// vybe-test: go/cover_runtime_testing/runtime_set_cpu_profile_rate
// origin: languages/go/tests/go/test_cover_runtime_testing.rs
// vybe-test-mode: compile

package main
import "runtime"
func main() { runtime.SetCPUProfileRate(100) }
