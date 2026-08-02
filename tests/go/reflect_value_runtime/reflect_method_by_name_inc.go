// vybe-test: go/reflect_value_runtime/reflect_method_by_name_inc
// origin: languages/go/tests/go/test_reflect_value_runtime.rs

package main
import "fmt"
import "reflect"
type Box struct { V int }
func (b *Box) Set(v int) { b.V = v }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { box := &Box{}
m := reflect.ValueOf(box).MethodByName("Set")
m.Call([]reflect.Value{reflect.ValueOf(15)})
__check(fmt.Sprint(box.V), "15") }
