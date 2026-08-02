// vybe-test: go/reflect_value_runtime/reflect_type_name_struct
// origin: languages/go/tests/go/test_reflect_value_runtime.rs

package main
import "fmt"
import "reflect"
type Widget struct{}
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { __check(fmt.Sprint(reflect.TypeOf(Widget{}).Name()), "Widget") }
