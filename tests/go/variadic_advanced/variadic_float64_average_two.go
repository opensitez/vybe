// vybe-test: go/variadic_advanced/variadic_float64_average_two
// origin: languages/go/tests/go/test_variadic_advanced.rs

package main
import "fmt"
func avg(nums ...float64) float64 { if len(nums) == 0 { return 0 }
s := 0.0
for _, n := range nums { s += n }
return s / float64(len(nums)) }
func main() { fmt.Println(avg(2.0, 4.0)) }
