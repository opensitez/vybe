// vybe-test: go/composite_literal_keys/map_nested_struct_and_slice_keys
// origin: languages/go/tests/go/test_composite_literal_keys.rs

package main
import "fmt"
type cell struct { n int }
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

func main() { m := map[string]struct { rows []cell }{ "t": {rows: []cell{{n: 1}, {n: 2}}} }
__p(fmt.Sprint(m["t"].rows[1].n))
__check("2")
}
