// vybe-test: go/reflect_value_runtime/reflect_call_method_changes_state
// origin: languages/go/tests/go/test_reflect_value_runtime.rs

package main
import "fmt"
import "reflect"
type Acc struct { Sum int }
func (a *Acc) Add(n int) { a.Sum += n }
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

func main() { acc := &Acc{}
reflect.ValueOf(acc).MethodByName("Add").Call([]reflect.Value{reflect.ValueOf(5)})
__p(fmt.Sprint(acc.Sum)) 
__check("5")
}
