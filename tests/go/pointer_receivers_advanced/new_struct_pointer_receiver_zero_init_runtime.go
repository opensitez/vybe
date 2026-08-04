// vybe-test: go/pointer_receivers_advanced/new_struct_pointer_receiver_zero_init_runtime
// origin: languages/go/tests/go/test_pointer_receivers_advanced.rs

package main
import "fmt"
type tally struct { sum int }
func (t *tally) add(v int) { t.sum += v }
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

func main() { value := new(tally)
value.add(3)
value.add(4)
__p(fmt.Sprint(value.sum))
__check("7")
}
