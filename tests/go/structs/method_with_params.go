// vybe-test: go/structs/method_with_params
// origin: languages/go/tests/go/test_structs.rs

package main
import "fmt"
type Wallet struct { Balance int } func (w Wallet) Withdraw(amount int) int { return w.Balance - amount } var __buf string

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

func main() { w := Wallet{Balance: 100}
__p(fmt.Sprint(w.Withdraw(30)))
__check("70")
}
