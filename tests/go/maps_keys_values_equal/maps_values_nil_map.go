// vybe-test: go/maps_keys_values_equal/maps_values_nil_map
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

func main() { var m map[string]int
vals := maps.Values(m)
__check(fmt.Sprint(vals == nil), "true")
__check(fmt.Sprint(len(vals)), "0") }
