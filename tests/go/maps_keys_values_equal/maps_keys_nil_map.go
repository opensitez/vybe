// vybe-test: go/maps_keys_values_equal/maps_keys_nil_map
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

func main() { var m map[int]bool
keys := maps.Keys(m)
__check(fmt.Sprint(keys == nil), "true")
__check(fmt.Sprint(len(keys)), "0") }
