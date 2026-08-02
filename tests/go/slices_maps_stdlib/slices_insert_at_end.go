// vybe-test: go/slices_maps_stdlib/slices_insert_at_end
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

func main() { s := []int{1,2}
s = slices.Insert(s, 2, 3)
__check(fmt.Sprint(len(s)), "3")
__check(fmt.Sprint(s[2]), "3") }
