// vybe-test: go/reflect_value_runtime/reflect_value_map_index
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

func main() { m := map[string]int{"a": 5}
v := reflect.ValueOf(m)
__check(fmt.Sprint(v.MapIndex(reflect.ValueOf("a")).Int()), "5") }
