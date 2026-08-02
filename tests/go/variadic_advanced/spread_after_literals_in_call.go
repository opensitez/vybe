// vybe-test: go/variadic_advanced/spread_after_literals_in_call
// origin: languages/go/tests/go/test_variadic_advanced.rs

package main
import "fmt"
func sum(nums ...int) int { t := 0
for _, n := range nums { t += n }
return t }
func main() { extra := []int{4, 5}
fmt.Println(sum(1, 2, 3, extra...)) }
