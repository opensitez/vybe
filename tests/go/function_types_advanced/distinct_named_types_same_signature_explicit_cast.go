// vybe-test: go/function_types_advanced/distinct_named_types_same_signature_explicit_cast
// origin: languages/go/tests/go/test_function_types_advanced.rs

package main
import "fmt"
type Adder func(int, int) int
type Combiner func(int, int) int
func use(c Combiner) int { return c(2, 5) }
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

func main() { var add Adder = func(a int, b int) int { return a + b }
__p(fmt.Sprint(use(Combiner(add)))) 
__check("7")
}
