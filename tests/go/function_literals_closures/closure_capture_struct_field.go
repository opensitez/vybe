// vybe-test: go/function_literals_closures/closure_capture_struct_field
// origin: languages/go/tests/go/test_function_literals_closures.rs

package main
import "fmt"
type counter struct { n int }
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

func main() { c := counter{n: 0}
bump := func() { c.n++ }
bump()
bump()
__p(fmt.Sprint(c.n)) 
__check("2")
}
