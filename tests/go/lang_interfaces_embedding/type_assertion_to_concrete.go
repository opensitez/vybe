// vybe-test: go/lang_interfaces_embedding/type_assertion_to_concrete
// origin: languages/go/tests/go/test_lang_interfaces_embedding.rs

package main
import "fmt"
type I interface { M() }
type T struct{}
func (T) M() {}
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

func main() { var i I = T{}
__p(fmt.Sprint(i.(T) == T{})) 
__check("true")
}
