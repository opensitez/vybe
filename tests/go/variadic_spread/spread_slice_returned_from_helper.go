// vybe-test: go/variadic_spread/spread_slice_returned_from_helper
// origin: languages/go/tests/go/test_variadic_spread.rs

package main
import "fmt"
func digits() []int { return []int{1, 2, 3} }
func sum(nums ...int) int { total := 0
for _, n := range nums { total += n }
return total }
func main() { fmt.Println(sum(digits()...))
}
