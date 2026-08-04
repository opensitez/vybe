// vybe-test: go/pointer_receivers_advanced/dual_receiver_value_read_pointer_write_runtime
// origin: languages/go/tests/go/test_pointer_receivers_advanced.rs

package main
import "fmt"
type ledger struct { balance int }
func (l ledger) snapshot() int { return l.balance }
func (l *ledger) deposit(v int) { l.balance += v }
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

func main() { value := ledger{balance: 10}
__p(fmt.Sprint(value.snapshot()))
value.deposit(5)
__p(fmt.Sprint(value.snapshot()))
__check("10\n15")
}
