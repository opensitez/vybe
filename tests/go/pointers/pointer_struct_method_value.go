// vybe-test: go/pointers/pointer_struct_method_value
// origin: languages/go/tests/go/test_pointers.rs

package main
import "fmt"
type Counter struct { N int }
func (c Counter) Inc() { c.N++ }
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

func main() { c := Counter{N: 0}
c.Inc()
__p(fmt.Sprint(c.N))
__check("0")
}
