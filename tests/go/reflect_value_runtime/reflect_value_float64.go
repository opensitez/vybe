// vybe-test: go/reflect_value_runtime/reflect_value_float64
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

func main() { v := reflect.ValueOf(2.5)
__check(fmt.Sprint(v.Float()), "2.5") }
