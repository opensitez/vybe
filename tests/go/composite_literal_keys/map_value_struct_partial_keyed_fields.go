// vybe-test: go/composite_literal_keys/map_value_struct_partial_keyed_fields
// origin: languages/go/tests/go/test_composite_literal_keys.rs

package main
import "fmt"
type point struct { x int
y int
label string }
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

func main() { m := map[string]point{"p": {y: 9, label: "home"}}
__p(fmt.Sprint(m["p"].y))
__p(fmt.Sprint(m["p"].x))
__p(fmt.Sprint(m["p"].label))
__check("9\n0\nhome")
}
