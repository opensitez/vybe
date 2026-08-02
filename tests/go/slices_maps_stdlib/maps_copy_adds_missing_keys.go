// vybe-test: go/slices_maps_stdlib/maps_copy_adds_missing_keys
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
src := map[string]int{"b": 2}
n := maps.Copy(dst, src)
__check(fmt.Sprint(n), "1")
__check(fmt.Sprint(dst["b"]), "2") }
