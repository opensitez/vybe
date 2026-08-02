// vybe-test: go/function_literals_closures/closure_with_variadic_params
// origin: languages/go/tests/go/test_function_literals_closures.rs

package main
import "fmt"
func main() { sum := func(nums ...int) int { total := 0
for _, n := range nums { total += n }
return total }
fmt.Println(sum(1, 2, 3)) }
