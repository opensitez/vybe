// vybe-test: go/nil_zero_semantics_extra/zero_value_array_field_runtime
// origin: languages/go/tests/go/test_nil_zero_semantics_extra.rs

package main
import "fmt"
type bag struct { values [2]int }
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

func main() { var b bag
__p(fmt.Sprint(b.values[1]))
__check("0")
}
