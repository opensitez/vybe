// vybe-test: go/methods_receivers_extra/method_with_array_field_runtime
// origin: languages/go/tests/go/test_methods_receivers_extra.rs

package main
import "fmt"
type bag struct { values [2]int }
func (b bag) second() int { return b.values[1] }
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

func main() { value := bag{values: [2]int{3, 9}}
__p(fmt.Sprint(value.second()))
__check("9")
}
