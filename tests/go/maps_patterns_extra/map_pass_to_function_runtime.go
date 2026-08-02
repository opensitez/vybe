// vybe-test: go/maps_patterns_extra/map_pass_to_function_runtime
// origin: languages/go/tests/go/test_maps_patterns_extra.rs

package main
import "fmt"
func total(values map[string]int) int { return values["a"] + values["b"] }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { __check(fmt.Sprint(total(map[string]int{"a": 2, "b": 4})), "6")
}
