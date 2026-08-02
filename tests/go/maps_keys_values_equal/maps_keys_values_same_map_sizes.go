// vybe-test: go/maps_keys_values_equal/maps_keys_values_same_map_sizes
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

func main() { m := map[int]int{1: 10, 2: 20, 3: 30, 4: 40}
__check(fmt.Sprint(len(maps.Keys(m))), "4")
__check(fmt.Sprint(len(maps.Values(m))), "4") }
