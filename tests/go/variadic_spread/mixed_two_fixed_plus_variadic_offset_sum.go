// vybe-test: go/variadic_spread/mixed_two_fixed_plus_variadic_offset_sum
// origin: languages/go/tests/go/test_variadic_spread.rs

package main
import "fmt"
func tally(base int, step int, nums ...int) int { total := base
for _, n := range nums { total += n + step }
return total }
func main() { fmt.Println(tally(100, 1, 2, 3))
}
