// vybe-test: go/cover_runtime_testing/quick_value
// origin: languages/go/tests/go/test_cover_runtime_testing.rs
// vybe-test-mode: compile

package main
import "testing/quick"
func main() { var x int
_ = quick.Value(&x, nil) }
