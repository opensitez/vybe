// vybe-test: go/generics_functions/generic_pair_swap
// origin: languages/go/tests/go/test_generics_functions.rs

package main
import "fmt"
func Swap[T any](a, b T) (T, T) { return b, a }
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

func main() { x, y := Swap(1, 2)
__p(fmt.Sprint(x))
__p(fmt.Sprint(y)) 
__check("2\n1")
}
