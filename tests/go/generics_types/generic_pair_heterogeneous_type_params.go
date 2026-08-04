// vybe-test: go/generics_types/generic_pair_heterogeneous_type_params
// origin: languages/go/tests/go/test_generics_types.rs

package main
import "fmt"
type Pair[A, B any] struct { First A
Second B }
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

func main() { p := Pair[int, string]{First: 9, Second: "go"}
__p(fmt.Sprint(p.First))
__p(fmt.Sprint(p.Second)) 
__check("9\ngo")
}
