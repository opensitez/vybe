// vybe-test: go/maps_keys_values_equal/maps_equal_func_both_nil
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

func main() { var a, b map[string]int
__check(fmt.Sprint(maps.EqualFunc(a, b, func(x, y int) bool { return true })), "true") }
