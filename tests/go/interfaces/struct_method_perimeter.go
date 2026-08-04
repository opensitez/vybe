// vybe-test: go/interfaces/struct_method_perimeter
// origin: languages/go/tests/go/test_interfaces.rs

package main
import "fmt"
type Rect struct { W int
H int } func (r Rect) Perimeter() int { return 2 * (r.W + r.H) } var __buf string

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

func main() { r := Rect{W: 3, H: 4}
__p(fmt.Sprint(r.Perimeter()))
__check("14")
}
