// vybe-test: go/method_values/pointer_receiver_method_value
// origin: languages/go/tests/go/test_method_values.rs

package main
import "fmt"
type acc struct { sum int }
func (a *acc) add(x int) { a.sum += x }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { a := &acc{}
inc := a.add
inc(4)
inc(5)
__check(fmt.Sprint(a.sum), "9") }
