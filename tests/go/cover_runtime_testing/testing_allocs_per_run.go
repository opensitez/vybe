// vybe-test: go/cover_runtime_testing/testing_allocs_per_run
// origin: languages/go/tests/go/test_cover_runtime_testing.rs
// vybe-test-mode: compile

package main
import "testing"
func main() { _ = testing.AllocsPerRun(1, func() { _ = make([]byte, 8) }) }
