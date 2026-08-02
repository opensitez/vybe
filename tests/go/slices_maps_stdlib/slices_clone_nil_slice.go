// vybe-test: go/slices_maps_stdlib/slices_clone_nil_slice
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
cp := slices.Clone(s)
__check(fmt.Sprint(cp == nil), "true")
__check(fmt.Sprint(len(cp)), "0") }
