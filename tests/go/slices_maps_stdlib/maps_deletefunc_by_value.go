// vybe-test: go/slices_maps_stdlib/maps_deletefunc_by_value
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

func main() { m := map[int]string{1: "keep", 2: "drop", 3: "drop"}
maps.DeleteFunc(m, func(k int, v string) bool { return v == "drop" })
__check(fmt.Sprint(len(m)), "1")
__check(fmt.Sprint(m[1]), "keep") }
