// vybe-test: go/maps_keys_values_equal/maps_equal_self_reference_semantics
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

func main() { m := map[int]int{1: 1}
__check(fmt.Sprint(maps.Equal(m, m)), "true") }
