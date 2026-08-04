// vybe-test: go/function_literals_closures/closure_filter_slice
// origin: languages/go/tests/go/test_function_literals_closures.rs

package main
import "fmt"
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

func main() { nums := []int{1, 2, 3, 4}
evens := func() []int { out := []int{}
for _, n := range nums { if n%2 == 0 { out = append(out, n) } }
return out }
r := evens()
__p(fmt.Sprint(len(r)))
__p(fmt.Sprint(r[0])) 
__check("2\n2")
}
