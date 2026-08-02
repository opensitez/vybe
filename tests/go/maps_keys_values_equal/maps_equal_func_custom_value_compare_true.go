// vybe-test: go/maps_keys_values_equal/maps_equal_func_custom_value_compare_true
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

func main() { a := map[int]int{1: 10, 2: 20}
b := map[int]int{1: 10, 2: 20}
__check(fmt.Sprint(maps.EqualFunc(a, b, func(x, y int) bool { return x == y })), "true") }
