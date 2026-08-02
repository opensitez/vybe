// vybe-test: go/variadic_spread/variadic_float64_sum
// origin: languages/go/tests/go/test_variadic_spread.rs

package main
import "fmt"
func sum(nums ...float64) float64 { total := 0.0
for _, n := range nums { total += n }
return total }
func main() { fmt.Println(sum(0.5, 1.5, 2.0))
}
