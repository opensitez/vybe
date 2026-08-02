// vybe-test: go/pointer_receivers_advanced/nil_pointer_receiver_nil_check_without_panic_runtime
// origin: languages/go/tests/go/test_pointer_receivers_advanced.rs

package main
import "fmt"
type node struct { id int }
func (n *node) absent() bool { return n == nil }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { var value *node
__check(fmt.Sprint(value.absent()), "true")
}
