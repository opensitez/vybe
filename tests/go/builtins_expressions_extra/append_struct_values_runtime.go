// vybe-test: go/builtins_expressions_extra/append_struct_values_runtime
// origin: languages/go/tests/go/test_builtins_expressions_extra.rs

package main
import "fmt"
type point struct { x int }
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

func main() { values := []point{}
values = append(values, point{x: 14})
__p(fmt.Sprint(values[0].x))
__check("14")
}
