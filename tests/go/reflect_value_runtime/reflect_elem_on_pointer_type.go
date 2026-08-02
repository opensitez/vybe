// vybe-test: go/reflect_value_runtime/reflect_elem_on_pointer_type
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

func main() { var x int
__check(fmt.Sprint(reflect.TypeOf(&x).Elem().Kind()), "int") }
