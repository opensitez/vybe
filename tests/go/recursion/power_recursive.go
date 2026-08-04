// vybe-test: go/recursion/power_recursive
// origin: languages/go/tests/go/test_recursion.rs

package main
import "fmt"
func pow(base int, exp int) int { if exp == 0 { return 1 }
return base * pow(base, exp-1) } var __buf string

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

func main() { __p(fmt.Sprint(pow(2, 10)))
__check("1024")
}
