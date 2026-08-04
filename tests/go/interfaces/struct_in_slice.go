// vybe-test: go/interfaces/struct_in_slice
// origin: languages/go/tests/go/test_interfaces.rs

package main
import "fmt"
type Point struct { X int
Y int } var __buf string

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

func main() { pts := []Point{{X: 1, Y: 2}, {X: 3, Y: 4}}
__p(fmt.Sprint(pts[0].X))
__p(fmt.Sprint(pts[1].Y))
__check("1\n4")
}
