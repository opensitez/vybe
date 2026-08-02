// vybe-test: go/slices_maps_stdlib/slices_insert_into_nil_slice
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

func main() { var s []int
s = slices.Insert(s, 0, 5, 6)
__check(fmt.Sprint(len(s)), "2")
__check(fmt.Sprint(s[1]), "6") }
