// vybe-test: go/defer_panic_recover_extra/defer_method_on_pointer_receiver_runtime
// origin: languages/go/tests/go/test_defer_panic_recover_extra.rs

package main
import "fmt"
type counter struct { n int }
func (c *counter) show() { __check(fmt.Sprint(c.n), "9") }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { value := &counter{n: 6}
defer value.show()
value.n = 9
}
