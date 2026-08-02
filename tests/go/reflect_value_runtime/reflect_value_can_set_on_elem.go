// vybe-test: go/reflect_value_runtime/reflect_value_can_set_on_elem
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

func main() { x := 1
v := reflect.ValueOf(&x).Elem()
__check(fmt.Sprint(v.CanSet()), "true") }
