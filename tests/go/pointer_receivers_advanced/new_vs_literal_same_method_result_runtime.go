// vybe-test: go/pointer_receivers_advanced/new_vs_literal_same_method_result_runtime
// origin: languages/go/tests/go/test_pointer_receivers_advanced.rs

package main
import "fmt"
type score struct { points int }
func (s *score) double() { s.points = s.points * 2 }
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

func main() { a := new(score)
a.points = 5
a.double()
b := &score{points: 5}
b.double()
__p(fmt.Sprint(a.points))
__p(fmt.Sprint(b.points))
__check("10\n10")
}
