// vybe-test: go/cover_runtime_testing/quick_check
// origin: languages/go/tests/go/test_cover_runtime_testing.rs
// vybe-test-mode: compile

package main
import "testing/quick"
func main() { f := func(x int) bool { return true }
_ = quick.Check(f, nil) }
