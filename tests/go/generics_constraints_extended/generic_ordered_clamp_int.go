// vybe-test: go/generics_constraints_extended/generic_ordered_clamp_int
// origin: languages/go/tests/go/test_generics_constraints_extended.rs

package main
import "fmt"
import "cmp"
func Clamp[T cmp.Ordered](v, lo, hi T) T { if cmp.Less(v, lo) { return lo }
if cmp.Less(hi, v) { return hi }
return v }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { __check(fmt.Sprint(Clamp(99, 0, 10)), "10") }
