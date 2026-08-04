// vybe-test: go/slices_maps_stdlib/maps_deletefunc_clears_all_entries
// origin: languages/go/tests/go/test_slices_maps_stdlib.rs

package main
import "fmt"
import "maps"
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

func main() { m := map[int]int{1: 1, 2: 2}
maps.DeleteFunc(m, func(k int, v int) bool { return true })
__p(fmt.Sprint(len(m))) 
__check("0")
}
