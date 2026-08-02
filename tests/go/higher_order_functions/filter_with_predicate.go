// vybe-test: go/higher_order_functions/filter_with_predicate
// origin: languages/go/tests/go/test_higher_order_functions.rs

package main
import "fmt"
func keep(nums []int, ok func(int) bool) []int { out := []int{}
for _, n := range nums { if ok(n) { out = append(out, n) } }
return out }
func main() { r := keep([]int{1,2,3,4}, func(n int) bool { return n%2 == 0 })
fmt.Println(len(r))
fmt.Println(r[0]) }
