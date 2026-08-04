// vybe-test: go/structs/struct_in_func_param
// origin: languages/go/tests/go/test_structs.rs

package main
import "fmt"
type Vec struct { X int
Y int } func dotProduct(a Vec, b Vec) int { return a.X*b.X + a.Y*b.Y } var __buf string

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

func main() { v1 := Vec{X: 2, Y: 3}
v2 := Vec{X: 4, Y: 5}
__p(fmt.Sprint(dotProduct(v1, v2)))
__check("23")
}
