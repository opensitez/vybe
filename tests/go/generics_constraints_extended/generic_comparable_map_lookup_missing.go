// vybe-test: go/generics_constraints_extended/generic_comparable_map_lookup_missing
// origin: languages/go/tests/go/test_generics_constraints_extended.rs

package main
import "fmt"
func Get[K comparable, V any](m map[K]V, k K) (V, bool) { v, ok := m[k]
return v, ok }
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

func main() { _, ok := Get(map[int]string{1: "a"}, 9)
__p(fmt.Sprint(ok)) 
__check("false")
}
