// vybe-test: go/higher_order_functions/callback_reduces_slice
// origin: languages/go/tests/go/test_higher_order_functions.rs

package main
import "fmt"
func fold(nums []int, combine func(int,int) int, init int) int { acc := init
for _, n := range nums { acc = combine(acc, n) }
return acc }
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

func main() { __p(fmt.Sprint(fold([]int{1,2,3}, func(a,b int) int { return a+b }, 0))) 
__check("6")
}
