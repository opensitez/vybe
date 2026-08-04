// vybe-test: go/functions_patterns_extra/tuple_return_ignored_second_runtime
// origin: languages/go/tests/go/test_functions_patterns_extra.rs

package main
import "fmt"
func pair() (int, string) { return 8, "unused" }
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

func main() { value, _ := pair()
__p(fmt.Sprint(value))
__check("8")
}
