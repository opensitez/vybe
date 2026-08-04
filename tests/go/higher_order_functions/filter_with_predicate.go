// vybe-test: go/higher_order_functions/filter_with_predicate
// origin: languages/go/tests/go/test_higher_order_functions.rs

package main
import "fmt"
func keep(nums []int, ok func(int) bool) []int { out := []int{}
for _, n := range nums { if ok(n) { out = append(out, n) } }
return out }
var __buf string

// __p appends one line, __pr appends without a newline.
func __p(s string) { __buf = __buf + s + "\n" }

func __pr(s string) { __buf = __buf + s }

// __check ends the program unless the collected output equals want. The final
// Println contributes a trailing newline the expected line vector never
// carried, so both forms are accepted.
func __check(want string) {
	if __buf != want && __buf != want+"\n" {
		fmt.Println("FAIL: want [" + want + "] got [" + __buf + "]")
		panic("assertion failed")
	}
}

func main() { r := keep([]int{1,2,3,4}, func(n int) bool { return n%2 == 0 })
__p(fmt.Sprint(len(r)))
__p(fmt.Sprint(r[0])) 
__check("2\n2")
}
