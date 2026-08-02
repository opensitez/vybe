// vybe-test: go/defer_lifo_extended/defer_method_receiver_evaluated_early
// origin: languages/go/tests/go/test_defer_lifo_extended.rs

package main
import "fmt"
type T struct { v int }
func (t T) show() { __check(fmt.Sprint(t.v), "1") }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { t := T{v: 1}
defer t.show()
t.v = 2
}
