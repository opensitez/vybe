// vybe-test: go/reflect_value_runtime/reflect_method_num_method
// origin: languages/go/tests/go/test_reflect_value_runtime.rs

package main
import "fmt"
import "reflect"
type T struct{}
func (T) A() {}
func (T) B() {}
func (*T) C() {}
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { __check(fmt.Sprint(reflect.TypeOf(T{}).NumMethod()), "2")
__check(fmt.Sprint(reflect.TypeOf(&T{}).NumMethod()), "3") }
