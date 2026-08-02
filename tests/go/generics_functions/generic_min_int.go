// vybe-test: go/generics_functions/generic_min_int
// origin: languages/go/tests/go/test_generics_functions.rs

package main
import "fmt"
func Min[T ~int](a, b T) T { if a < b { return a }
return b }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { __check(fmt.Sprint(Min(3, 7)), "3") }
