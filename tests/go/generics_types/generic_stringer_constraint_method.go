// vybe-test: go/generics_types/generic_stringer_constraint_method
// origin: languages/go/tests/go/test_generics_types.rs

package main
import "fmt"
type Stringer interface { String() string }
type Show[T Stringer] struct{}
func (Show[T]) Display(v T) string { return v.String() }
type Tag struct { Label string }
func (t Tag) String() string { return t.Label }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { __check(fmt.Sprint(Show[Tag]{}.Display(Tag{Label: "ok"})), "ok") }
