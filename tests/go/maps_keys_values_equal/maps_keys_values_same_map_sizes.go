// vybe-test: go/maps_keys_values_equal/maps_keys_values_same_map_sizes
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

func main() { m := map[int]int{1: 10, 2: 20, 3: 30, 4: 40}
__p(fmt.Sprint(len(maps.Keys(m))))
__p(fmt.Sprint(len(maps.Values(m)))) 
__check("4\n4")
}
