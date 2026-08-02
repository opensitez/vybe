// vybe-test: go/pointer_receivers_advanced/address_of_field_then_pointer_method_runtime
// origin: languages/go/tests/go/test_pointer_receivers_advanced.rs

package main
import "fmt"
type meter struct { reading int }
func (m *meter) set(v int) { m.reading = v }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { value := meter{reading: 0}
fieldPtr := &value.reading
*fieldPtr = 3
value.set(7)
__check(fmt.Sprint(value.reading), "7")
}
