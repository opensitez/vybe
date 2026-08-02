// vybe-test: go/function_types_advanced/named_reducer_type_two_param_callback
// origin: languages/go/tests/go/test_function_types_advanced.rs

package main
import "fmt"
type Reducer func(int, int) int
func fold(values []int, r Reducer, init int) int { acc := init
for _, v := range values { acc = r(acc, v) }
return acc }
func main() { fmt.Println(fold([]int{1, 2, 3}, Reducer(func(a int, b int) int { return a + b }), 0)) }
