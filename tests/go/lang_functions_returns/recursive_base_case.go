// vybe-test: go/lang_functions_returns/recursive_base_case
// origin: languages/go/tests/go/test_lang_functions_returns.rs

package main
import "fmt"
func fact(n int) int { if n <= 1 { return 1 }
return n * fact(n-1) }
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

func main() { __p(fmt.Sprint(fact(4))) 
__check("24")
}
