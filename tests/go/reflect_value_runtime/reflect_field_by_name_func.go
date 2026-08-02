// vybe-test: go/reflect_value_runtime/reflect_field_by_name_func
// origin: languages/go/tests/go/test_reflect_value_runtime.rs

package main
import "fmt"
import "reflect"
type Row struct { Alpha int
Beta int
Gamma int }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { f, ok := reflect.TypeOf(Row{}).FieldByNameFunc(func(name string) bool { return len(name) == 5 })
__check(fmt.Sprint(ok), "true")
__check(fmt.Sprint(f.Name), "Alpha") }
