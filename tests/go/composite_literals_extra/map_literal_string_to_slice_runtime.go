// vybe-test: go/composite_literals_extra/map_literal_string_to_slice_runtime
// origin: languages/go/tests/go/test_composite_literals_extra.rs

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

func main() { values := map[string][]int{"odd": {1, 3, 5}}
__p(fmt.Sprint(values["odd"][2]))
__check("5")
}
