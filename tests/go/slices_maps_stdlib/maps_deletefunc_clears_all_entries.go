// vybe-test: go/slices_maps_stdlib/maps_deletefunc_clears_all_entries
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

func main() { m := map[int]int{1: 1, 2: 2}
maps.DeleteFunc(m, func(k int, v int) bool { return true })
__check(fmt.Sprint(len(m)), "0") }
