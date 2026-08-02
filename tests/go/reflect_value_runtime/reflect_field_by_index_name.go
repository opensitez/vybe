// vybe-test: go/reflect_value_runtime/reflect_field_by_index_name
// origin: languages/go/tests/go/test_reflect_value_runtime.rs

package main
import "fmt"
import "reflect"
type Pair struct { A int
B string }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { f := reflect.TypeOf(Pair{}).Field(0)
__check(fmt.Sprint(f.Name), "A") }
