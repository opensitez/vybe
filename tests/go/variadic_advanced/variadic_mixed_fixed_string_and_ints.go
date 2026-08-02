// vybe-test: go/variadic_advanced/variadic_mixed_fixed_string_and_ints
// origin: languages/go/tests/go/test_variadic_advanced.rs

package main
import "fmt"
func tagSum(label string, nums ...int) int { t := len(label)
for _, n := range nums { t += n }
return t }
func main() { fmt.Println(tagSum("go", 1, 2)) }
