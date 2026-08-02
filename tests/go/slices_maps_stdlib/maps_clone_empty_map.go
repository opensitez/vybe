// vybe-test: go/slices_maps_stdlib/maps_clone_empty_map
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

func main() { cp := maps.Clone(map[string]int{})
__check(fmt.Sprint(len(cp)), "0") }
