// vybe-test: go/slice_aliasing_extra/slice_pass_to_function_compile
// origin: languages/go/tests/go/test_slice_aliasing_extra.rs
// vybe-test-mode: compile

package main
func use(values []int) int { return len(values) }
func main() { _ = use([]int{1}) }
