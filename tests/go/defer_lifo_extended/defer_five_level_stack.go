// vybe-test: go/defer_lifo_extended/defer_five_level_stack
// origin: languages/go/tests/go/test_defer_lifo_extended.rs

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

func main() { defer __p(fmt.Sprint("a"))
defer __p(fmt.Sprint("b"))
defer __p(fmt.Sprint("c"))
defer __p(fmt.Sprint("d"))
defer __p(fmt.Sprint("e"))
__check("e\nd\nc\nb\na")
}
