// vybe-test: go/slices_maps_advanced/map_delete_non_existent
// origin: languages/go/tests/go/test_slices_maps_advanced.rs

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

func main() { m := map[int]int{1: 1}
delete(m, 2)
__p(fmt.Sprint(len(m)))
__check("1")
}
