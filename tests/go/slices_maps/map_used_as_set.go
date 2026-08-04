// vybe-test: go/slices_maps/map_used_as_set
// origin: languages/go/tests/go/test_slices_maps.rs

package main
import "fmt"
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

func main() { s := map[int]bool{1: true, 2: true, 3: true}
__p(fmt.Sprint(s[2]))
__p(fmt.Sprint(s[5]))
__check("true\nfalse")
}
