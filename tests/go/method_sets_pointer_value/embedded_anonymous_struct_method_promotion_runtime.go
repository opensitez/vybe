// vybe-test: go/method_sets_pointer_value/embedded_anonymous_struct_method_promotion_runtime
// origin: languages/go/tests/go/test_method_sets_pointer_value.rs

package main
import "fmt"
type coords struct { x int
y int }
func (c coords) sum() int { return c.x + c.y }
type point struct { coords }
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

func main() { p := point{coords: coords{x: 2, y: 5}}
__p(fmt.Sprint(p.sum())) 
__check("7")
}
