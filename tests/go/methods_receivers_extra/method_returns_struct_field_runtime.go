// vybe-test: go/methods_receivers_extra/method_returns_struct_field_runtime
// origin: languages/go/tests/go/test_methods_receivers_extra.rs

package main
import "fmt"
type point struct { x int }
func (p point) value() int { return p.x }
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

func main() { __p(fmt.Sprint(point{x: 12}.value()))
__check("12")
}
