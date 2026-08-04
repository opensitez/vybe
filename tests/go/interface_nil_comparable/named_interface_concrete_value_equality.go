// vybe-test: go/interface_nil_comparable/named_interface_concrete_value_equality
// origin: languages/go/tests/go/test_interface_nil_comparable.rs

package main
import "fmt"
type counter interface { count() int }
type tally struct { n int }
func (t tally) count() int { return t.n }
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

func main() { var left counter = tally{n: 5}
var right counter = tally{n: 5}
__p(fmt.Sprint(left == right)) 
__check("true")
}
