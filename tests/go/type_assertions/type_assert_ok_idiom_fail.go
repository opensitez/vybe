// vybe-test: go/type_assertions/type_assert_ok_idiom_fail
// origin: languages/go/tests/go/test_type_assertions.rs

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

func main() { var x interface{} = 42
s, ok := x.(string)
__p(fmt.Sprint(s))
__p(fmt.Sprint(ok))
__check("\nfalse")
}
