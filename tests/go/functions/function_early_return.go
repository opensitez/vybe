// vybe-test: go/functions/function_early_return
// origin: languages/go/tests/go/test_functions.rs

package main
import "fmt"
func check(n int) string { if n < 0 { return "neg" }
return "ok" } var __buf string

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

func main() { __p(fmt.Sprint(check(-1)))
__p(fmt.Sprint(check(1)))
__check("neg\nok")
}
