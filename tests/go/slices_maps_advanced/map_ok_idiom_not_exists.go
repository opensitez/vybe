// vybe-test: go/slices_maps_advanced/map_ok_idiom_not_exists
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

func main() { m := map[string]int{"a": 1}
v, ok := m["b"]
__p(fmt.Sprint(v))
__p(fmt.Sprint(ok))
__check("0\nfalse")
}
