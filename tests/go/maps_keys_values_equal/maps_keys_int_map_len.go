// vybe-test: go/maps_keys_values_equal/maps_keys_int_map_len
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

func main() { m := map[int]string{1: "a", 2: "b", 3: "c"}
keys := maps.Keys(m)
__check(fmt.Sprint(len(keys)), "3") }
