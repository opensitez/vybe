// vybe-test: go/variadic_advanced/spread_subslice_middle_two
// origin: languages/go/tests/go/test_variadic_advanced.rs

package main
import "fmt"
func sum(nums ...int) int { t := 0
for _, n := range nums { t += n }
return t }
func main() { all := []int{5, 1, 2, 3, 9}
fmt.Println(sum(all[1:3]...)) }
