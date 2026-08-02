// vybe-test: go/slices_maps_stdlib/slices_grow_zero_len_slice
// origin: languages/go/tests/go/test_slices_maps_stdlib.rs

package main
import "fmt"
import "slices"
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { s := make([]int, 0, 1)
s = slices.Grow(s, 2)
__check(fmt.Sprint(len(s)), "0")
__check(fmt.Sprint(cap(s) >= 2), "true") }
