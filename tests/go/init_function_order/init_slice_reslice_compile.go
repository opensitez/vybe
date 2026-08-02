// vybe-test: go/init_function_order/init_slice_reslice_compile
// origin: languages/go/tests/go/test_init_function_order.rs
// vybe-test-mode: compile

package main
var part []int
func init() { all := []int{1, 2, 3, 4}
part = all[1:3] }
func main() { _ = part[0] }
