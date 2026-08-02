// vybe-test: go/cover_runtime_testing/runtime_set_mutex_profile_fraction
// origin: languages/go/tests/go/test_cover_runtime_testing.rs
// vybe-test-mode: compile

package main
import "runtime"
func main() { runtime.SetMutexProfileFraction(1) }
