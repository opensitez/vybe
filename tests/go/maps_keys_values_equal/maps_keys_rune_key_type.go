// vybe-test: go/maps_keys_values_equal/maps_keys_rune_key_type
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

func main() { m := map[rune]int{'a': 1, 'b': 2}
__check(fmt.Sprint(len(maps.Keys(m))), "2") }
