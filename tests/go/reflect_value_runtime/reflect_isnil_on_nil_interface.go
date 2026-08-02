// vybe-test: go/reflect_value_runtime/reflect_isnil_on_nil_interface
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

func main() { var i interface{}
__check(fmt.Sprint(reflect.ValueOf(i).IsNil()), "true") }
