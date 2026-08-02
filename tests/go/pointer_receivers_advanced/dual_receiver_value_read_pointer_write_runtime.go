// vybe-test: go/pointer_receivers_advanced/dual_receiver_value_read_pointer_write_runtime
// origin: languages/go/tests/go/test_pointer_receivers_advanced.rs

package main
import "fmt"
type ledger struct { balance int }
func (l ledger) snapshot() int { return l.balance }
func (l *ledger) deposit(v int) { l.balance += v }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { value := ledger{balance: 10}
__check(fmt.Sprint(value.snapshot()), "10")
value.deposit(5)
__check(fmt.Sprint(value.snapshot()), "15")
}
