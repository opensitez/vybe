// vybe-test: go/lang_functions_returns/higher_order_map
// origin: languages/go/tests/go/test_lang_functions_returns.rs

package main
import "fmt"
func mapInts(xs []int, f func(int) int) []int { out := make([]int, len(xs))
for i, v := range xs { out[i] = f(v) }
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

func main() { __p(fmt.Sprint(mapInts([]int{1,2}, func(x int) int { return x*2 })[1])) 
__check("4")
}
