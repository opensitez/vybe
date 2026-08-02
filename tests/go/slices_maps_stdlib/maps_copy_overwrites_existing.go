// vybe-test: go/slices_maps_stdlib/maps_copy_overwrites_existing
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

func main() { dst := map[string]int{"a": 1}
src := map[string]int{"a": 9}
maps.Copy(dst, src)
__check(fmt.Sprint(dst["a"]), "9") }
