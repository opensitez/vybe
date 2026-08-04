// vybe-test: go/generics_constraints_extended/generic_tilde_struct_slice_len
// origin: languages/go/tests/go/test_generics_constraints_extended.rs

package main
import "fmt"
type Ints []int
func Total[S ~[]int](s S) int { sum := 0
for _, v := range s { sum += v }
return sum }
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

func main() { __p(fmt.Sprint(Total(Ints{1, 2, 3}))) 
__check("6")
}
