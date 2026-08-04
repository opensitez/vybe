// vybe-test: go/slices_maps_stdlib/slices_delete_middle_range
// origin: languages/go/tests/go/test_slices_maps_stdlib.rs

package main
import "fmt"
import "slices"
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

func main() { s := []int{0,1,2,3,4}
s = slices.Delete(s, 1, 3)
__p(fmt.Sprint(len(s)))
__p(fmt.Sprint(s[0]))
__p(fmt.Sprint(s[1]))
__p(fmt.Sprint(s[2])) 
__check("3\n0\n3\n4")
}
