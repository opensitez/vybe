// vybe-test: go/variadic_spread/variadic_int_minimum
// origin: languages/go/tests/go/test_variadic_spread.rs

package main
import "fmt"
func minimum(nums ...int) int { m := nums[0]
for _, n := range nums { if n < m { m = n } }
return m }
func main() { fmt.Println(minimum(5, 1, 8, 2))
}
