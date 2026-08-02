// vybe-test: go/reflect_value_runtime/reflect_valueof_string_interface
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

func main() { v := reflect.ValueOf("vybe")
__check(fmt.Sprint(v.String()), "vybe") }
