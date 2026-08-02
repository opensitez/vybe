// vybe-test: go/function_types_advanced/pointer_receiver_run_with_func_param
// origin: languages/go/tests/go/test_function_types_advanced.rs

package main
import "fmt"
type acc struct { total int }
func (a *acc) addEach(values []int, combine func(int, int) int) { for _, v := range values { a.total = combine(a.total, v) } }
func main() { value := acc{}
value.addEach([]int{1, 2, 3}, func(a int, b int) int { return a + b })
fmt.Println(value.total) }
