// vybe-test: go/slices_maps_stdlib/slices_compact_all_identical
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

func main() { s := []int{7,7,7,7}
s = slices.Compact(s)
__check(fmt.Sprint(len(s)), "1")
__check(fmt.Sprint(s[0]), "7") }
