// vybe-test: go/maps_keys_values_equal/maps_values_after_delete
// origin: languages/go/tests/go/test_maps_keys_values_equal.rs

package main
import "fmt"
import "maps"
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { m := map[int]int{1: 1, 2: 2, 3: 3}
delete(m, 2)
__check(fmt.Sprint(len(maps.Values(m))), "2") }
