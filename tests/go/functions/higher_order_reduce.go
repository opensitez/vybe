// vybe-test: go/functions/higher_order_reduce
// origin: languages/go/tests/go/test_functions.rs

package main
import "fmt"
func reduce(s []int, init int, f func(int, int) int) int { acc := init
for _, v := range s { acc = f(acc, v) }
return acc } var __buf string

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

func main() { total := reduce([]int{1,2,3,4,5}, 0, func(a int, b int) int { return a + b })
__p(fmt.Sprint(total))
__check("15")
}
