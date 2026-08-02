// vybe-test: go/pointer_receivers_advanced/value_method_then_pointer_method_sequence_runtime
// origin: languages/go/tests/go/test_pointer_receivers_advanced.rs

package main
import "fmt"
type account struct { balance int }
func (a account) funds() int { return a.balance }
func (a *account) credit(v int) { a.balance += v }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { value := account{balance: 20}
__check(fmt.Sprint(value.funds()), "20")
value.credit(7)
__check(fmt.Sprint(value.funds()), "27")
}
