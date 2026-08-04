// vybe-test: go/generics_constraints_extended/generic_ordered_clamp_int
// origin: languages/go/tests/go/test_generics_constraints_extended.rs

package main
import "fmt"
import "cmp"
func Clamp[T cmp.Ordered](v, lo, hi T) T { if cmp.Less(v, lo) { return lo }
if cmp.Less(hi, v) { return hi }
return v }
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

func main() { __p(fmt.Sprint(Clamp(99, 0, 10))) 
__check("10")
}
