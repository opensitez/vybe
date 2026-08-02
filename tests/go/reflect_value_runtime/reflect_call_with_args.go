// vybe-test: go/reflect_value_runtime/reflect_call_with_args
// origin: languages/go/tests/go/test_reflect_value_runtime.rs

package main
import "fmt"
import "reflect"
func Add(a, b int) int { return a + b }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { fv := reflect.ValueOf(Add)
out := fv.Call([]reflect.Value{reflect.ValueOf(3), reflect.ValueOf(4)})
__check(fmt.Sprint(out[0].Int()), "7") }
