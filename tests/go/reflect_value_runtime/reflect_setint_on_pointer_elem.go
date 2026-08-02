// vybe-test: go/reflect_value_runtime/reflect_setint_on_pointer_elem
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

func main() { x := 0
v := reflect.ValueOf(&x).Elem()
v.SetInt(42)
__check(fmt.Sprint(x), "42") }
