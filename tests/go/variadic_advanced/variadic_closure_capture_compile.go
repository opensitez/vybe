// vybe-test: go/variadic_advanced/variadic_closure_capture_compile
// origin: languages/go/tests/go/test_variadic_advanced.rs
// vybe-test-mode: compile

package main
func main() { base := 1
fn := func(nums ...int) int { return base + len(nums) }
_ = fn(2, 3) }
