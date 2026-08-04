// vybe-test: go/method_sets_pointer_value/dual_receivers_value_read_pointer_write_runtime
// origin: languages/go/tests/go/test_method_sets_pointer_value.rs

package main
import "fmt"
type wallet struct { cash int }
func (w wallet) balance() int { return w.cash }
func (w *wallet) deposit(v int) { w.cash += v }
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

func main() { w := wallet{cash: 10}
__p(fmt.Sprint(w.balance()))
w.deposit(5)
__p(fmt.Sprint(w.balance())) 
__check("10\n15")
}
