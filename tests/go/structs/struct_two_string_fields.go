// vybe-test: go/structs/struct_two_string_fields
// origin: languages/go/tests/go/test_structs.rs

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

func main() { p := Point{X: 3, Y: 4}
__p(fmt.Sprint(p.X))
__p(fmt.Sprint(p.Y))
__check("3\n4")
}
