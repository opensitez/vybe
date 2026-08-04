// vybe-test: go/generics_constraints_extended/generic_comparable_interface_key
// origin: languages/go/tests/go/test_generics_constraints_extended.rs

package main
import "fmt"
func Count[K comparable, V any](m map[K]V) int { return len(m) }
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

func main() { type I interface { ~int }
__p(fmt.Sprint(Count(map[int]string{1: "a"}))) 
__check("1")
}
