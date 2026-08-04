// vybe-test: go/generics_constraints_extended/generic_comparable_map_equal_keys
// origin: languages/go/tests/go/test_generics_constraints_extended.rs

package main
import "fmt"
func SameKeys[K comparable, V any](a, b map[K]V) bool { if len(a) != len(b) { return false }
for k := range a { if _, ok := b[k]; !ok { return false } }
return true }
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

func main() { a := map[string]int{"x": 1}
b := map[string]int{"x": 2}
__p(fmt.Sprint(SameKeys(a, b))) 
__check("true")
}
