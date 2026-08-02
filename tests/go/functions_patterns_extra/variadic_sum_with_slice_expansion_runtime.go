// vybe-test: go/functions_patterns_extra/variadic_sum_with_slice_expansion_runtime
// origin: languages/go/tests/go/test_functions_patterns_extra.rs

package main
import "fmt"
func sum(values ...int) int { total := 0
for _, v := range values { total += v }
return total }
func main() { nums := []int{4, 5, 6}
fmt.Println(sum(nums...))
}
