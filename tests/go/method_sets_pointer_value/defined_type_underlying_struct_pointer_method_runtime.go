// vybe-test: go/method_sets_pointer_value/defined_type_underlying_struct_pointer_method_runtime
// origin: languages/go/tests/go/test_method_sets_pointer_value.rs

package main
import "fmt"
type meters float64
func (m *meters) scale(f float64) { *m = meters(float64(*m) * f) }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { var m meters = 100
m.scale(2)
__check(fmt.Sprint(float64(m)), "200") }
