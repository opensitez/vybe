// vybe-test: go/slices_maps_stdlib/slices_insert_into_nil_slice
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

func main() { var s []int
s = slices.Insert(s, 0, 5, 6)
__p(fmt.Sprint(len(s)))
__p(fmt.Sprint(s[1])) 
__check("2\n6")
}
