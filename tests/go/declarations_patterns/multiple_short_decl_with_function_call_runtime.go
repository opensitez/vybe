// vybe-test: go/declarations_patterns/multiple_short_decl_with_function_call_runtime
// origin: languages/go/tests/go/test_declarations_patterns.rs

package main
import "fmt"
func pair() (int, int) { return 8, 9 }
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

func main() { a, b := pair()
__p(fmt.Sprint(a + b))
__check("17")
}
