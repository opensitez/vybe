// vybe-test: go/slices_maps_stdlib/maps_deletefunc_by_key_prefix
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

func main() { m := map[string]int{"pre1": 1, "pre2": 2, "other": 3}
maps.DeleteFunc(m, func(k string, v int) bool { return len(k) >= 5 })
__check(fmt.Sprint(len(m)), "2")
__check(fmt.Sprint(m["pre1"]), "1") }
