// vybe-test: go/functions/higher_order_filter
// origin: languages/go/tests/go/test_functions.rs

package main
import "fmt"
func filter(s []int, f func(int) bool) []int { r := []int{}
for _, v := range s { if f(v) { r = append(r, v) } }
return r } var __buf string

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

func main() { evens := filter([]int{1,2,3,4,5}, func(x int) bool { return x % 2 == 0 })
for _, v := range evens { __p(fmt.Sprint(v))
} 
__check("2\n4")
}
