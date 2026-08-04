// vybe-test: go/maps_keys_values_equal/maps_equal_identical_maps
// origin: languages/go/tests/go/test_maps_keys_values_equal.rs

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

func main() { a := map[string]int{"x": 1, "y": 2}
b := map[string]int{"x": 1, "y": 2}
__p(fmt.Sprint(maps.Equal(a, b))) 
__check("true")
}
