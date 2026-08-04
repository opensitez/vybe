// vybe-test: go/generics_functions/generic_min_int
// origin: languages/go/tests/go/test_generics_functions.rs

package main
import "fmt"
func Min[T ~int](a, b T) T { if a < b { return a }
return b }
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

func main() { __p(fmt.Sprint(Min(3, 7))) 
__check("3")
}
