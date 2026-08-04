// vybe-test: go/function_types_advanced/struct_field_func_cast_to_named_type
// origin: languages/go/tests/go/test_function_types_advanced.rs

package main
import "fmt"
type Mapper func(int) int
type holder struct { fn func(int) int }
func apply(v int, m Mapper) int { return m(v) }
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

func main() { value := holder{fn: func(v int) int { return v + 2 }}
__p(fmt.Sprint(apply(5, Mapper(value.fn)))) 
__check("7")
}
