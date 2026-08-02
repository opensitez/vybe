// vybe-test: go/methods_receivers_extra/pointer_receiver_called_twice_runtime
// origin: languages/go/tests/go/test_methods_receivers_extra.rs

package main
import "fmt"
type counter struct { n int }
func (c *counter) add(v int) { c.n += v }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { value := counter{}
value.add(2)
value.add(5)
__check(fmt.Sprint(value.n), "7")
}
