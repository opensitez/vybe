// vybe-test: go/variadic_spread/forward_variadic_single_delegate
// origin: languages/go/tests/go/test_variadic_spread.rs

package main
import "fmt"
func sink(nums ...int) int { total := 0
for _, n := range nums { total += n }
return total }
func relay(nums ...int) int { return sink(nums...) }
func main() { fmt.Println(relay(4, 5))
}
