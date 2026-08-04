// vybe-test: go/defer_lifo_extended/defer_with_recover_then_next_defer
// origin: languages/go/tests/go/test_defer_lifo_extended.rs

package main
import "fmt"
func run() { defer __p(fmt.Sprint("last"))
defer func() { recover() }()
defer __p(fmt.Sprint("mid"))
panic("p") }
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

func main() { run() 
__check("last\nmid")
}
