// vybe-test: go/generics_functions/generic_slice_len
// origin: languages/go/tests/go/test_generics_functions.rs

package main
import "fmt"
func Len[T any](s []T) int { return len(s) }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { __check(fmt.Sprint(Len([]int{1,2,3})), "3") }
