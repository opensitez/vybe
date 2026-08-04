// vybe-test: go/method_values/method_expression_requires_receiver
// origin: languages/go/tests/go/test_method_values.rs

package main
import "fmt"
type box struct { v int }
func (b box) get() int { return b.v }
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

func main() { f := box.get
__p(fmt.Sprint(f(box{v:9}))) 
__check("9")
}
