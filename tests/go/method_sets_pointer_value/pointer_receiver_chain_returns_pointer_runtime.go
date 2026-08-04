// vybe-test: go/method_sets_pointer_value/pointer_receiver_chain_returns_pointer_runtime
// origin: languages/go/tests/go/test_method_sets_pointer_value.rs

package main
import "fmt"
type node struct { val int }
func (n *node) add(v int) *node { n.val += v
return n }
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

func main() { n := &node{val: 1}
n.add(2).add(3)
__p(fmt.Sprint(n.val)) 
__check("6")
}
