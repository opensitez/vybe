// vybe-test: go/maps_keys_values_equal/maps_equal_same_key_different_value
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

func main() { a := map[string]int{"k": 5}
b := map[string]int{"k": 6}
__check(fmt.Sprint(maps.Equal(a, b)), "false") }
