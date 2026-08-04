// vybe-test: go/generics_types/generic_triple_three_type_parameters
// origin: languages/go/tests/go/test_generics_types.rs

package main
import "fmt"
type Triple[A, B, C any] struct { A A
B B
C C }
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

func main() { t := Triple[int, string, bool]{A: 1, B: "x", C: true}
__p(fmt.Sprint(t.A))
__p(fmt.Sprint(t.B))
__p(fmt.Sprint(t.C)) 
__check("1\nx\ntrue")
}
