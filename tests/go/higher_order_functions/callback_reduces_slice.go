// vybe-test: go/higher_order_functions/callback_reduces_slice
// origin: languages/go/tests/go/test_higher_order_functions.rs

package main
import "fmt"
func fold(nums []int, combine func(int,int) int, init int) int { acc := init
for _, n := range nums { acc = combine(acc, n) }
return acc }
func main() { fmt.Println(fold([]int{1,2,3}, func(a,b int) int { return a+b }, 0)) }
