// vybe-test: go/method_sets_pointer_value/defined_type_underlying_struct_value_method_runtime
// origin: languages/go/tests/go/test_method_sets_pointer_value.rs

package main
import "fmt"
type meters float64
func (m meters) km() float64 { return float64(m) / 1000 }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { __check(fmt.Sprint(meters(2500).km()), "2.5") }
