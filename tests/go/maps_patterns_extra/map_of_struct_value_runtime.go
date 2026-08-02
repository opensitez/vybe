// vybe-test: go/maps_patterns_extra/map_of_struct_value_runtime
// origin: languages/go/tests/go/test_maps_patterns_extra.rs

package main
import "fmt"
type point struct { x int }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { values := map[string]point{"a": {x: 13}}
__check(fmt.Sprint(values["a"].x), "13")
}
