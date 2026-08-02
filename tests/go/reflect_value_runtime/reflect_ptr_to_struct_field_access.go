// vybe-test: go/reflect_value_runtime/reflect_ptr_to_struct_field_access
// origin: languages/go/tests/go/test_reflect_value_runtime.rs

package main
import "fmt"
import "reflect"
type Data struct { N int }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { d := &Data{N: 8}
v := reflect.ValueOf(d).Elem().Field(0)
__check(fmt.Sprint(v.Int()), "8") }
