// vybe-test: go/reflect_value_runtime/reflect_setstring_on_pointer_elem
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

func main() { s := ""
v := reflect.ValueOf(&s).Elem()
v.SetString("hello")
__check(fmt.Sprint(s), "hello") }
