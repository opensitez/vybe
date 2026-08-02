// vybe-test: go/slices_maps_stdlib/slices_contains_int_found
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

func main() { __check(fmt.Sprint(slices.Contains([]int{1,2,3}, 2)), "true") }
