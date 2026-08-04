// vybe-test: go/lang_expressions/function_call_args_eval_left_to_right
// origin: languages/go/tests/go/test_lang_expressions.rs

package main
import "fmt"
func f(a, b int) int { return a*10+b }
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

func main() { i := 0
i++
__p(fmt.Sprint(f(i, i))) 
__check("12")
}
