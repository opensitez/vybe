// vybe-test: go/pointer_receivers_advanced/address_of_field_then_pointer_method_runtime
// origin: languages/go/tests/go/test_pointer_receivers_advanced.rs

package main
import "fmt"
type meter struct { reading int }
func (m *meter) set(v int) { m.reading = v }
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

func main() { value := meter{reading: 0}
fieldPtr := &value.reading
*fieldPtr = 3
value.set(7)
__p(fmt.Sprint(value.reading))
__check("7")
}
