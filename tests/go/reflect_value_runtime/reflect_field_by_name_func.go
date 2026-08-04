// vybe-test: go/reflect_value_runtime/reflect_field_by_name_func
// origin: languages/go/tests/go/test_reflect_value_runtime.rs

package main
import "fmt"
import "reflect"
type Row struct { Alpha int
Beta int
Gamma int }
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

func main() { f, ok := reflect.TypeOf(Row{}).FieldByNameFunc(func(name string) bool { return len(name) == 5 })
__p(fmt.Sprint(ok))
__p(fmt.Sprint(f.Name)) 
__check("true\nAlpha")
}
