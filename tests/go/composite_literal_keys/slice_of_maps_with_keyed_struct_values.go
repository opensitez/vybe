// vybe-test: go/composite_literal_keys/slice_of_maps_with_keyed_struct_values
// origin: languages/go/tests/go/test_composite_literal_keys.rs

package main
import "fmt"
type pair struct { a int
b int }
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

func main() { s := []map[string]pair{{"x": {b: 2, a: 1}}, {"y": pair{a: 3, b: 4}}}
__p(fmt.Sprint(s[0]["x"].a))
__p(fmt.Sprint(s[1]["y"].b))
__check("1\n4")
}
