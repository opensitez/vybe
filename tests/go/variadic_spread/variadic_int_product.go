// vybe-test: go/variadic_spread/variadic_int_product
// origin: languages/go/tests/go/test_variadic_spread.rs

package main
import "fmt"
func product(nums ...int) int { p := 1
for _, n := range nums { p *= n }
return p }
func main() { fmt.Println(product(2, 3, 4))
}
