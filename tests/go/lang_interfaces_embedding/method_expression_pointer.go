// vybe-test: go/lang_interfaces_embedding/method_expression_pointer
// origin: languages/go/tests/go/test_lang_interfaces_embedding.rs

package main
import "fmt"
type N int
func (n *N) Inc() { *n++ }
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

func main() { var v N = 1
f := (*N).Inc
f(&v)
__p(fmt.Sprint(v)) 
__check("2")
}
