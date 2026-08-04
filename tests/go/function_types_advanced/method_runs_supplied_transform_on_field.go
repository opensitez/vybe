// vybe-test: go/function_types_advanced/method_runs_supplied_transform_on_field
// origin: languages/go/tests/go/test_function_types_advanced.rs

package main
import "fmt"
type gauge struct { value int }
func (g *gauge) mapValue(mapper func(int) int) { g.value = mapper(g.value) }
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

func main() { g := gauge{value: 4}
g.mapValue(func(v int) int { return v * 2 })
__p(fmt.Sprint(g.value)) 
__check("8")
}
