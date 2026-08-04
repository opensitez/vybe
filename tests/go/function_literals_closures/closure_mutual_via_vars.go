// vybe-test: go/function_literals_closures/closure_mutual_via_vars
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

func main() { var even func(int) bool
var odd func(int) bool
even = func(n int) bool { if n == 0 { return true }
return odd(n-1) }
odd = func(n int) bool { if n == 0 { return false }
return even(n-1) }
__p(fmt.Sprint(even(4)))
__p(fmt.Sprint(odd(3))) 
__check("true\ntrue")
}
