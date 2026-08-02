// vybe-test: go/maps_patterns_extra/map_string_concat_values_runtime
// origin: languages/go/tests/go/test_maps_patterns_extra.rs

package main
import "fmt"
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { values := map[string]string{"a": "vy", "b": "be"}
__check(fmt.Sprint(values["a"] + values["b"]), "vybe")
}
