// vybe-test: go/cover_runtime_testing/quick_config
// origin: languages/go/tests/go/test_cover_runtime_testing.rs
// vybe-test-mode: compile

package main
import "testing/quick"
func main() { cfg := &quick.Config{MaxCount: 10}
f := func(x int) bool { return true }
_ = quick.Check(f, cfg) }
