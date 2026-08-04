// vybe-test: go/functions/recursive_function
// origin: languages/go/tests/go/test_functions.rs

package main
import "fmt"
func factorial(n int) int { if n <= 1 { return 1 }
return n * factorial(n-1) } var __buf string

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

func main() { __p(fmt.Sprint(factorial(5)))
__check("120")
}
