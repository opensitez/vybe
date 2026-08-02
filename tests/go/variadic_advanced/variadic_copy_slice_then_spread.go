// vybe-test: go/variadic_advanced/variadic_copy_slice_then_spread
// origin: languages/go/tests/go/test_variadic_advanced.rs

package main
import "fmt"
func sum(nums ...int) int { t := 0
for _, n := range nums { t += n }
return t }
func main() { src := []int{1, 2}
dup := append([]int(nil), src...)
fmt.Println(sum(dup...)) }
