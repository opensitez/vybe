// vybe-test: go/functions_patterns_extra/function_literal_returns_struct_runtime
// origin: languages/go/tests/go/test_functions_patterns_extra.rs

package main
import "fmt"
type pair struct { a int
b int }
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

func main() { build := func() pair { return pair{a: 3, b: 4} }
value := build()
__p(fmt.Sprint(value.a + value.b))
__check("7")
}
