// vybe-test: go/reflect_value_runtime/reflect_callslice_variadic
// origin: languages/go/tests/go/test_reflect_value_runtime.rs

package main
import "fmt"
import "reflect"
func Sum(nums ...int) int { s := 0
for _, n := range nums { s += n }
return s }
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

func main() { fv := reflect.ValueOf(Sum)
out := fv.CallSlice([]reflect.Value{reflect.ValueOf(1), reflect.ValueOf(2), reflect.ValueOf(3)})
__p(fmt.Sprint(out[0].Int())) 
__check("6")
}
