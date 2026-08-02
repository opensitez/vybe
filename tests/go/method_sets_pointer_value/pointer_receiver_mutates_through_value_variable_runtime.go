// vybe-test: go/method_sets_pointer_value/pointer_receiver_mutates_through_value_variable_runtime
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

func main() { v := cell{n: 2}
v.inc()
__check(fmt.Sprint(v.n), "3") }
