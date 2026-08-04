// vybe-test: go/recursion/fibonacci_base_zero
// origin: languages/go/tests/go/test_recursion.rs

package main
import "fmt"
func fib(n int) int { if n <= 0 { return 0 }
if n == 1 { return 1 }
return fib(n-1) + fib(n-2) } var __buf string

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

func main() { __p(fmt.Sprint(fib(0)))
__check("0")
}
