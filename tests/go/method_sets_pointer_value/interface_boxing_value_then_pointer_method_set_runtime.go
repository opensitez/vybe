// vybe-test: go/method_sets_pointer_value/interface_boxing_value_then_pointer_method_set_runtime
// origin: languages/go/tests/go/test_method_sets_pointer_value.rs

package main
import "fmt"
type grower interface { grow() }
type plant struct { h int }
func (p *plant) grow() { p.h++ }
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

func main() { p := &plant{h: 1}
var g grower = p
g.grow()
__p(fmt.Sprint(p.h)) 
__check("2")
}
