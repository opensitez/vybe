// vybe-test: go/reflect_value_runtime/reflect_interface_roundtrip_string
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

func main() { v := reflect.ValueOf("go")
s := v.Interface().(string)
__check(fmt.Sprint(s), "go") }
