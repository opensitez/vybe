// vybe-test: go/reflect_value_runtime/reflect_numfield_struct
// origin: languages/go/tests/go/test_reflect_value_runtime.rs

package main
import "fmt"
import "reflect"
type Person struct { Name string
Age int }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { __check(fmt.Sprint(reflect.TypeOf(Person{}).NumField()), "2") }
