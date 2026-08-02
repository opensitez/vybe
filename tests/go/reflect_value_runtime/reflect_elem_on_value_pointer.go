// vybe-test: go/reflect_value_runtime/reflect_elem_on_value_pointer
// origin: languages/go/tests/go/test_reflect_value_runtime.rs

package main
import "fmt"
import "reflect"
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { x := 10
v := reflect.ValueOf(&x)
__check(fmt.Sprint(v.Elem().Int()), "10") }
