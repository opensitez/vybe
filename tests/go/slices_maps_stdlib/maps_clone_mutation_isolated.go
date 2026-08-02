// vybe-test: go/slices_maps_stdlib/maps_clone_mutation_isolated
// origin: languages/go/tests/go/test_slices_maps_stdlib.rs

package main
import "fmt"
import "maps"
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { orig := map[string]int{"a": 1}
cp := maps.Clone(orig)
cp["a"] = 9
__check(fmt.Sprint(orig["a"]), "1")
__check(fmt.Sprint(cp["a"]), "9") }
