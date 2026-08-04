// vybe-test: go/reflect_value_runtime/reflect_call_with_args
// origin: languages/go/tests/go/test_reflect_value_runtime.rs

package main
import "fmt"
import "reflect"
func Add(a, b int) int { return a + b }
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

func main() { fv := reflect.ValueOf(Add)
out := fv.Call([]reflect.Value{reflect.ValueOf(3), reflect.ValueOf(4)})
__p(fmt.Sprint(out[0].Int())) 
__check("7")
}
