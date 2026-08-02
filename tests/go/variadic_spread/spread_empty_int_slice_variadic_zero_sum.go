// vybe-test: go/variadic_spread/spread_empty_int_slice_variadic_zero_sum
// origin: languages/go/tests/go/test_variadic_spread.rs

package main
import "fmt"
func sum(nums ...int) int { total := 0
for _, n := range nums { total += n }
return total }
func main() { fmt.Println(sum([]int{}...))
}
