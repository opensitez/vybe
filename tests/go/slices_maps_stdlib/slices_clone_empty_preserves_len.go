// vybe-test: go/slices_maps_stdlib/slices_clone_empty_preserves_len
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

func main() { cp := slices.Clone([]int{})
__check(fmt.Sprint(len(cp)), "0") }
