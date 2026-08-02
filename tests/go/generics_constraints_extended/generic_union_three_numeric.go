// vybe-test: go/generics_constraints_extended/generic_union_three_numeric
// origin: languages/go/tests/go/test_generics_constraints_extended.rs

package main
import "fmt"
func Sign[T int | int64 | float64](v T) int { if v < 0 { return -1 }
if v > 0 { return 1 }
return 0 }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { __check(fmt.Sprint(Sign(int64(-3))), "-1") }
