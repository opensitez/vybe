// vybe-test: go/variadic_spread/forward_variadic_injects_count_before_delegate
// origin: languages/go/tests/go/test_variadic_spread.rs

package main
import "fmt"
func emit(nums ...int) { for _, n := range nums { fmt.Println(n) } }
func relay(nums ...int) { fmt.Println(len(nums))
emit(nums...) }
func main() { relay(7, 8)
}
