// vybe-test: go/methods/method_on_custom_string
// origin: languages/go/tests/go/test_methods.rs

package main
import "fmt"
type MyStr string
func (s MyStr) Len() int { return len(s) }
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

func main() { var s MyStr = "hello"
__p(fmt.Sprint(s.Len()))
__check("5")
}
