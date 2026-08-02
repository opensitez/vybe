// vybe-test: go/methods_receivers_extra/pointer_method_on_new_runtime
// origin: languages/go/tests/go/test_methods_receivers_extra.rs

package main
import "fmt"
type counter struct { n int }
func (c *counter) bump() { c.n++ }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { value := new(counter)
value.bump()
__check(fmt.Sprint(value.n), "1")
}
