// vybe-test: go/lang_functions_returns/variadic_slice_spread
// origin: languages/go/tests/go/test_lang_functions_returns.rs

package main
import "fmt"
func sum(xs ...int) int { t := 0
for _, x := range xs { t += x }
return t }
func main() { xs := []int{1,2}
fmt.Println(sum(xs...)) }
