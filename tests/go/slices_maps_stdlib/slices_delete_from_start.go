// vybe-test: go/slices_maps_stdlib/slices_delete_from_start
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

func main() { s := []int{9,8,7}
s = slices.Delete(s, 0, 1)
__check(fmt.Sprint(len(s)), "2")
__check(fmt.Sprint(s[0]), "8") }
