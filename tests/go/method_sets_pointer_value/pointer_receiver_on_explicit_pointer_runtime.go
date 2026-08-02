// vybe-test: go/method_sets_pointer_value/pointer_receiver_on_explicit_pointer_runtime
// origin: languages/go/tests/go/test_method_sets_pointer_value.rs

package main
import "fmt"
type cell struct { n int }
func (c *cell) inc() { c.n++ }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { p := &cell{n: 2}
p.inc()
__check(fmt.Sprint(p.n), "3") }
