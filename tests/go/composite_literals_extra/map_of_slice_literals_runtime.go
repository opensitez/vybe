// vybe-test: go/composite_literals_extra/map_of_slice_literals_runtime
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

func main() { values := map[string][]string{"langs": {"go", "rust"}}
__p(fmt.Sprint(values["langs"][0]))
__check("go")
}
