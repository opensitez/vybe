// vybe-test: go/init_function_order/init_writes_to_map_of_slices_compile
// origin: languages/go/tests/go/test_init_function_order.rs
// vybe-test-mode: compile

package main
var groups = map[string][]int{}
func init() { groups["a"] = []int{1} }
func init() { groups["a"] = append(groups["a"], 2) }
func main() { _ = groups }
