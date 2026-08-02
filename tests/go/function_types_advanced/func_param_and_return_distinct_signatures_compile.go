// vybe-test: go/function_types_advanced/func_param_and_return_distinct_signatures_compile
// origin: languages/go/tests/go/test_function_types_advanced.rs
// vybe-test-mode: compile

package main
func read(fn func() int) int { return fn() }
func write(fn func(int)) {}
func main() { write(func(v int) {})
_ = read(func() int { return 1 }) }
