// vybe-test: go/method_sets_pointer_value/pointer_receiver_chain_returns_pointer_runtime
// origin: languages/go/tests/go/test_method_sets_pointer_value.rs

package main
import "fmt"
type node struct { val int }
func (n *node) add(v int) *node { n.val += v
return n }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { n := &node{val: 1}
n.add(2).add(3)
__check(fmt.Sprint(n.val), "6") }
