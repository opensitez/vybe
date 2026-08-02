// vybe-test: go/slices_maps_stdlib/slices_grow_expands_capacity
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

func main() { s := make([]int, 2, 2)
s[0], s[1] = 1, 2
s = slices.Grow(s, 3)
__check(fmt.Sprint(len(s)), "2")
__check(fmt.Sprint(cap(s) >= 5), "true")
__check(fmt.Sprint(s[0]), "1") }
