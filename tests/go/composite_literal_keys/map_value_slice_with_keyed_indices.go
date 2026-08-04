// vybe-test: go/composite_literal_keys/map_value_slice_with_keyed_indices
// origin: languages/go/tests/go/test_composite_literal_keys.rs

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

func main() { m := map[string][]int{"data": {0: 5, 2: 7}}
__p(fmt.Sprint(len(m["data"])))
__p(fmt.Sprint(m["data"][0]))
__p(fmt.Sprint(m["data"][2]))
__check("3\n5\n7")
}
