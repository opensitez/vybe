// vybe-test: go/defer/defer_nested_func
// origin: languages/go/tests/go/test_defer.rs

package main
import "fmt"
func inner() { defer __p(fmt.Sprint("inner def"))
__p(fmt.Sprint("inner run"))
}
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

func main() { defer __p(fmt.Sprint("main def"))
inner()
__p(fmt.Sprint("main run"))
__check("inner run\ninner def\nmain run\nmain def")
}
