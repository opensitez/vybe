// vybe-test: go/maps_keys_values_equal/maps_equal_func_abs_value_equivalence
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

func main() { a := map[int]int{1: -5}
b := map[int]int{1: 5}
eq := maps.EqualFunc(a, b, func(x, y int) bool { if x < 0 { x = -x }; if y < 0 { y = -y }; return x == y })
__check(fmt.Sprint(eq), "true") }
