// vybe-test: go/variadic_advanced/variadic_func_value_call_compile
// origin: languages/go/tests/go/test_variadic_advanced.rs
// vybe-test-mode: compile

package main
func main() { fn := func(nums ...int) int { return len(nums) }
_ = fn(1, 2) }
