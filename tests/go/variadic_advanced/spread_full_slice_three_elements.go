// vybe-test: go/variadic_advanced/spread_full_slice_three_elements
// origin: languages/go/tests/go/test_variadic_advanced.rs

package main
import "fmt"
func sum(nums ...int) int { t := 0
for _, n := range nums { t += n }
return t }
func main() { batch := []int{1, 2, 3}
fmt.Println(sum(batch...)) }
