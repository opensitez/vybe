// vybe-test: go/reflect_value_runtime/reflect_ptr_to_struct_field_access
// origin: languages/go/tests/go/test_reflect_value_runtime.rs

package main
import "fmt"
import "reflect"
type Data struct { N int }
var __buf string

// __p appends one line, __pr appends without a newline.
func __p(s string) { __buf = __buf + s + "\n" }

func __pr(s string) { __buf = __buf + s }

// __check ends the program unless the collected output equals want. The final
// Println contributes a trailing newline the expected line vector never
// carried, so both forms are accepted.
func __check(want string) {
	if __buf != want && __buf != want+"\n" {
		fmt.Println("FAIL: want [" + want + "] got [" + __buf + "]")
		panic("assertion failed")
	}
}

func main() { d := &Data{N: 8}
v := reflect.ValueOf(d).Elem().Field(0)
__p(fmt.Sprint(v.Int())) 
__check("8")
}
