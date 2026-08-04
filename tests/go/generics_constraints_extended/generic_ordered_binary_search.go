// vybe-test: go/generics_constraints_extended/generic_ordered_binary_search
// origin: languages/go/tests/go/test_generics_constraints_extended.rs

package main
import "fmt"
import "cmp"
import "slices"
func Find[T cmp.Ordered](s []T, target T) (int, bool) { return slices.BinarySearch(s, target) }
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

func main() { i, ok := Find([]int{1, 3, 5}, 3)
__p(fmt.Sprint(i))
__p(fmt.Sprint(ok)) 
__check("1\ntrue")
}
