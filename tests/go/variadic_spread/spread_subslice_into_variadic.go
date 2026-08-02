// vybe-test: go/variadic_spread/spread_subslice_into_variadic
// origin: languages/go/tests/go/test_variadic_spread.rs

package main
import "fmt"
func sum(nums ...int) int { total := 0
for _, n := range nums { total += n }
return total }
func main() { all := []int{10, 1, 2, 3}
fmt.Println(sum(all[1:]...))
}
