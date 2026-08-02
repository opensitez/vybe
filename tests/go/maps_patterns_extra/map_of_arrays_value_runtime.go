// vybe-test: go/maps_patterns_extra/map_of_arrays_value_runtime
// origin: languages/go/tests/go/test_maps_patterns_extra.rs

package main
import "fmt"
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { values := map[string][2]int{"a": [2]int{3, 4}}
__check(fmt.Sprint(values["a"][1]), "4")
}
