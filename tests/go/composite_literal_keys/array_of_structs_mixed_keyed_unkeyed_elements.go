// vybe-test: go/composite_literal_keys/array_of_structs_mixed_keyed_unkeyed_elements
// origin: languages/go/tests/go/test_composite_literal_keys.rs

package main
import "fmt"
type pair struct { left int
right int }
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

func main() { a := [3]pair{{right: 2, left: 1}, {3, 4}, pair{left: 5, right: 6}}
__p(fmt.Sprint(a[0].left))
__p(fmt.Sprint(a[1].right))
__p(fmt.Sprint(a[2].left))
__check("1\n4\n5")
}
