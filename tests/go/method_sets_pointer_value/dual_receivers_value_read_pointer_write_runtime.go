// vybe-test: go/method_sets_pointer_value/dual_receivers_value_read_pointer_write_runtime
// origin: languages/go/tests/go/test_method_sets_pointer_value.rs

package main
import "fmt"
type wallet struct { cash int }
func (w wallet) balance() int { return w.cash }
func (w *wallet) deposit(v int) { w.cash += v }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { w := wallet{cash: 10}
__check(fmt.Sprint(w.balance()), "10")
w.deposit(5)
__check(fmt.Sprint(w.balance()), "15") }
