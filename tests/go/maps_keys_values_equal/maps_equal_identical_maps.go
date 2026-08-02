// vybe-test: go/maps_keys_values_equal/maps_equal_identical_maps
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

func main() { a := map[string]int{"x": 1, "y": 2}
b := map[string]int{"x": 1, "y": 2}
__check(fmt.Sprint(maps.Equal(a, b)), "true") }
