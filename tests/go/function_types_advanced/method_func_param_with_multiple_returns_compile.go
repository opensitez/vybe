// vybe-test: go/function_types_advanced/method_func_param_with_multiple_returns_compile
// origin: languages/go/tests/go/test_function_types_advanced.rs
// vybe-test-mode: compile

package main
type Splitter func(int) (int, int)
type divider struct{}
func (divider) parts(v int, split Splitter) (int, int) { return split(v) }
func main() { _, _ = divider{}.parts(9, Splitter(func(v int) (int, int) { return v / 2, v % 2 })) }
