// vybe-test: go/maps_keys_values_equal/maps_equal_one_nil_other_nonempty
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

func main() { var a map[int]int
b := map[int]int{1: 1}
__check(fmt.Sprint(maps.Equal(a, b)), "false") }
