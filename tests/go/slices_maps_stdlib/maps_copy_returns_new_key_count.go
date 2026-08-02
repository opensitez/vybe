// vybe-test: go/slices_maps_stdlib/maps_copy_returns_new_key_count
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

func main() { dst := map[string]int{"x": 1}
src := map[string]int{"x": 2, "y": 3}
n := maps.Copy(dst, src)
__check(fmt.Sprint(n), "1")
__check(fmt.Sprint(len(dst)), "2") }
