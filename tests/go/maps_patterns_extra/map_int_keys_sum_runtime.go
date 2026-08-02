// vybe-test: go/maps_patterns_extra/map_int_keys_sum_runtime
// origin: languages/go/tests/go/test_maps_patterns_extra.rs

package main
import "fmt"
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { values := map[int]int{1: 2, 2: 3}
__check(fmt.Sprint(values[1] + values[2]), "5")
}
