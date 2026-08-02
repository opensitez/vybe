// vybe-test: go/reflect_value_runtime/reflect_field_by_name_found
// origin: languages/go/tests/go/test_reflect_value_runtime.rs

package main
import "fmt"
import "reflect"
type S struct { Score int }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { f, ok := reflect.TypeOf(S{}).FieldByName("Score")
__check(fmt.Sprint(ok), "true")
__check(fmt.Sprint(f.Name), "Score") }
