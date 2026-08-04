// vybe-test: go/function_types_advanced/struct_func_field_reassigned_and_called
// origin: languages/go/tests/go/test_function_types_advanced.rs

package main
import "fmt"
type holder struct { fn func(string) string }
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

func main() { value := holder{}
__p(fmt.Sprint(value.fn == nil))
value.fn = func(s string) string { return s + "!" }
__p(fmt.Sprint(value.fn("go"))) 
__check("true\ngo!")
}
