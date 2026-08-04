// vybe-test: go/struct_embedding_extra/struct_return_from_function_runtime
// origin: languages/go/tests/go/test_struct_embedding_extra.rs

package main
import "fmt"
type point struct { x int
y int }
func build() point { return point{x: 4, y: 6} }
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

func main() { value := build()
__p(fmt.Sprint(value.x))
__p(fmt.Sprint(value.y))
__check("4\n6")
}
