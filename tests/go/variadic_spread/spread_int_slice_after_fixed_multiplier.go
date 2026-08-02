// vybe-test: go/variadic_spread/spread_int_slice_after_fixed_multiplier
// origin: languages/go/tests/go/test_variadic_spread.rs

package main
import "fmt"
func scale(factor int, nums ...int) int { total := 0
for _, n := range nums { total += n * factor }
return total }
func main() { batch := []int{2, 3, 4}
fmt.Println(scale(10, batch...))
}
