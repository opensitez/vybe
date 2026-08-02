// vybe-test: go/maps_keys_values_equal/maps_values_int_map_len
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

func main() { m := map[string]int{"a": 10, "b": 20, "c": 30}
vals := maps.Values(m)
__check(fmt.Sprint(len(vals)), "3") }
