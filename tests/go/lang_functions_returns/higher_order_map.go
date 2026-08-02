// vybe-test: go/lang_functions_returns/higher_order_map
// origin: languages/go/tests/go/test_lang_functions_returns.rs

package main
import "fmt"
func mapInts(xs []int, f func(int) int) []int { out := make([]int, len(xs))
for i, v := range xs { out[i] = f(v) }
return out }
func main() { fmt.Println(mapInts([]int{1,2}, func(x int) int { return x*2 })[1]) }
