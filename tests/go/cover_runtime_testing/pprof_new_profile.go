// vybe-test: go/cover_runtime_testing/pprof_new_profile
// origin: languages/go/tests/go/test_cover_runtime_testing.rs
// vybe-test-mode: compile

package main
import "runtime/pprof"
func main() { _ = pprof.NewProfile("custom") }
