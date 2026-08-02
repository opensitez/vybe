// vybe-test: go/higher_order_functions/higher_order_generic_callback
// origin: languages/go/tests/go/test_higher_order_functions.rs
// vybe-test-mode: compile

package main
func mapInts(src []int, f func(int) int) []int { out := make([]int, len(src))
for i, v := range src { out[i] = f(v) }
return out }
func main() { _ = mapInts([]int{1}, func(x int) int { return x }) }
