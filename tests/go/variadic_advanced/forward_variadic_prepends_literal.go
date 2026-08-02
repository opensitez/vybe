// vybe-test: go/variadic_advanced/forward_variadic_prepends_literal
// origin: languages/go/tests/go/test_variadic_advanced.rs

package main
import "fmt"
func emit(nums ...int) { for _, n := range nums { fmt.Println(n) } }
func relay(nums ...int) { emit(append([]int{0}, nums...)...) }
func main() { relay(5, 6) }
