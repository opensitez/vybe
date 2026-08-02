// vybe-test: go/reflect_value_runtime/reflect_call_method_changes_state
// origin: languages/go/tests/go/test_reflect_value_runtime.rs

package main
import "fmt"
import "reflect"
type Acc struct { Sum int }
func (a *Acc) Add(n int) { a.Sum += n }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { acc := &Acc{}
reflect.ValueOf(acc).MethodByName("Add").Call([]reflect.Value{reflect.ValueOf(5)})
__check(fmt.Sprint(acc.Sum), "5") }
