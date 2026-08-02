// vybe-test: go/slices_maps_stdlib/slices_clone_mutation_isolated
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

func main() { orig := []int{1,2,3}
cp := slices.Clone(orig)
cp[0] = 99
__check(fmt.Sprint(orig[0]), "1")
__check(fmt.Sprint(cp[0]), "99") }
