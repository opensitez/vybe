// vybe-test: go/method_sets_pointer_value/embedded_value_type_pointer_method_on_outer_value_runtime
// origin: languages/go/tests/go/test_method_sets_pointer_value.rs

package main
import "fmt"
type engine struct { rpm int }
func (e *engine) rev() { e.rpm++ }
type car struct { engine }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { c := car{engine: engine{rpm: 1000}}
c.rev()
__check(fmt.Sprint(c.rpm), "1001") }
