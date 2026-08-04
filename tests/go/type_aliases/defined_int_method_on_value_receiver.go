// vybe-test: go/type_aliases/defined_int_method_on_value_receiver
// origin: languages/go/tests/go/test_type_aliases.rs

package main
import "fmt"
type Meters int
func (m Meters) Double() int { return int(m) * 2 }
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

func main() { __p(fmt.Sprint(Meters(5).Double())) 
__check("10")
}
