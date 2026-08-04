// vybe-test: go/generics_constraints_extended/generic_any_slice_reverse
// origin: languages/go/tests/go/test_generics_constraints_extended.rs

package main
import "fmt"
func Reverse[T any](s []T) { for i, j := 0, len(s)-1; i < j; i, j = i+1, j-1 { s[i], s[j] = s[j], s[i] } }
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

func main() { a := []int{1, 2, 3}
Reverse(a)
__p(fmt.Sprint(a[0]))
__p(fmt.Sprint(a[2])) 
__check("3\n1")
}
