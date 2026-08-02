// vybe-test: go/generics_functions/generic_pair_swap
// origin: languages/go/tests/go/test_generics_functions.rs

package main
import "fmt"
func Swap[T any](a, b T) (T, T) { return b, a }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { x, y := Swap(1, 2)
__check(fmt.Sprint(x), "2")
__check(fmt.Sprint(y), "1") }
