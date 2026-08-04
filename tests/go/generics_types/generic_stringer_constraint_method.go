// vybe-test: go/generics_types/generic_stringer_constraint_method
// origin: languages/go/tests/go/test_generics_types.rs

package main
import "fmt"
type Stringer interface { String() string }
type Show[T Stringer] struct{}
func (Show[T]) Display(v T) string { return v.String() }
type Tag struct { Label string }
func (t Tag) String() string { return t.Label }
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

func main() { __p(fmt.Sprint(Show[Tag]{}.Display(Tag{Label: "ok"}))) 
__check("ok")
}
