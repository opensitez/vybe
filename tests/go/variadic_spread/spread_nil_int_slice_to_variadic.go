// vybe-test: go/variadic_spread/spread_nil_int_slice_to_variadic
// origin: languages/go/tests/go/test_variadic_spread.rs
// vybe-test-mode: compile

package main
import "fmt"
func sum(nums ...int) int { return len(nums) }
func main() { var s []int
fmt.Println(sum(s...)) }
