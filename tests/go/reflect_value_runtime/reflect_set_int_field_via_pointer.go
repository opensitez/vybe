// vybe-test: go/reflect_value_runtime/reflect_set_int_field_via_pointer
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

func main() { d := &Data{}
f := reflect.ValueOf(d).Elem().Field(0)
f.SetInt(33)
__check(fmt.Sprint(d.N), "33") }
