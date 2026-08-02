// vybe-test: go/variadic_spread/spread_single_element_int_slice
// origin: languages/go/tests/go/test_variadic_spread.rs

package main
import "fmt"
func sum(nums ...int) int { total := 0
for _, n := range nums { total += n }
return total }
func main() { fmt.Println(sum([]int{9}...))
}
