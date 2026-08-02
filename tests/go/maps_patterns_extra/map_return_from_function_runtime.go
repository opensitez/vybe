// vybe-test: go/maps_patterns_extra/map_return_from_function_runtime
// origin: languages/go/tests/go/test_maps_patterns_extra.rs

package main
import "fmt"
func build() map[string]int { return map[string]int{"a": 6} }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { __check(fmt.Sprint(build()["a"]), "6")
}
