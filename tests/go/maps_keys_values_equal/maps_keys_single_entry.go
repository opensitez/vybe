// vybe-test: go/maps_keys_values_equal/maps_keys_single_entry
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

func main() { m := map[int]int{42: 99}
keys := maps.Keys(m)
__check(fmt.Sprint(len(keys)), "1")
__check(fmt.Sprint(keys[0]), "42") }
