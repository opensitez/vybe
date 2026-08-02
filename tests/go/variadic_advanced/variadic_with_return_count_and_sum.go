// vybe-test: go/variadic_advanced/variadic_with_return_count_and_sum
// origin: languages/go/tests/go/test_variadic_advanced.rs

package main
import "fmt"
func stats(nums ...int) (int, int) { t := 0
for _, n := range nums { t += n }
return len(nums), t }
func main() { c, s := stats(2, 3, 4)
fmt.Println(c)
fmt.Println(s) }
