// vybe-test: go/variadic_advanced/mixed_two_fixed_int_variadic_sum
// origin: languages/go/tests/go/test_variadic_advanced.rs

package main
import "fmt"
func offset(base int, step int, vals ...int) int { t := base
for _, v := range vals { t += v + step }
return t }
func main() { fmt.Println(offset(10, 1, 2, 3)) }
