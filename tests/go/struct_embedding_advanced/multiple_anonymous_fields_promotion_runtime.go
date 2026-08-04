// vybe-test: go/struct_embedding_advanced/multiple_anonymous_fields_promotion_runtime
// origin: languages/go/tests/go/test_struct_embedding_advanced.rs

package main
import "fmt"
type axis struct { x int }
type ord struct { y int }
type point struct { axis
ord }
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

func main() { value := point{axis: axis{x: 4}, ord: ord{y: 6}}
__p(fmt.Sprint(value.x + value.y))
__check("10")
}
