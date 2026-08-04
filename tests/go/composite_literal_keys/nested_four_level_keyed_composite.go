// vybe-test: go/composite_literal_keys/nested_four_level_keyed_composite
// origin: languages/go/tests/go/test_composite_literal_keys.rs

package main
import "fmt"
type leaf struct { v int }
type branch struct { leaves []leaf }
type tree struct { parts []branch }
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

func main() { tr := tree{parts: []branch{{leaves: []leaf{{v: 99}}}}}
__p(fmt.Sprint(tr.parts[0].leaves[0].v))
__check("99")
}
