// vybe-test: go/method_values/pointer_receiver_method_value
// origin: languages/go/tests/go/test_method_values.rs

package main
import "fmt"
type acc struct { sum int }
func (a *acc) add(x int) { a.sum += x }
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

func main() { a := &acc{}
inc := a.add
inc(4)
inc(5)
__p(fmt.Sprint(a.sum)) 
__check("9")
}
