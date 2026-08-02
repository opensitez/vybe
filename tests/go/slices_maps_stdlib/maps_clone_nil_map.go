// vybe-test: go/slices_maps_stdlib/maps_clone_nil_map
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

func main() { var m map[string]int
cp := maps.Clone(m)
__check(fmt.Sprint(cp == nil), "true")
__check(fmt.Sprint(len(cp)), "0") }
