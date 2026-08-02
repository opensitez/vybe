// vybe-test: go/composite_literal_keys/slice_of_arrays_with_keyed_inner_indices
// origin: languages/go/tests/go/test_composite_literal_keys.rs

package main
import "fmt"
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { s := [][2]int{{0: 5, 1: 6}, {1: 8, 0: 7}}
__check(fmt.Sprint(s[0][0]), "5")
__check(fmt.Sprint(s[1][1]), "8")
}
