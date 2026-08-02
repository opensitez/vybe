// vybe-test: go/reflect_value_runtime/reflect_setbool_on_pointer_elem
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

func main() { b := false
v := reflect.ValueOf(&b).Elem()
v.SetBool(true)
__check(fmt.Sprint(b), "true") }
