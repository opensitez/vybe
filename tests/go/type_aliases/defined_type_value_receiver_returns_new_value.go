// vybe-test: go/type_aliases/defined_type_value_receiver_returns_new_value
// origin: languages/go/tests/go/test_type_aliases.rs

package main
import "fmt"
type Offset int
func (o Offset) next() Offset { return o + 1 }
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

func main() { __p(fmt.Sprint(Offset(2).next())) 
__check("3")
}
