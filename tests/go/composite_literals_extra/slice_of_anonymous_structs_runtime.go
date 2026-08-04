// vybe-test: go/composite_literals_extra/slice_of_anonymous_structs_runtime
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

func main() { values := []struct { n int }{{n: 3}, {n: 6}}
__p(fmt.Sprint(values[0].n + values[1].n))
__check("9")
}
