// vybe-test: go/pointer_receivers_advanced/value_method_then_pointer_method_sequence_runtime
// origin: languages/go/tests/go/test_pointer_receivers_advanced.rs

package main
import "fmt"
type account struct { balance int }
func (a account) funds() int { return a.balance }
func (a *account) credit(v int) { a.balance += v }
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

func main() { value := account{balance: 20}
__p(fmt.Sprint(value.funds()))
value.credit(7)
__p(fmt.Sprint(value.funds()))
__check("20\n27")
}
