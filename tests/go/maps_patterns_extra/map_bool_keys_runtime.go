// vybe-test: go/maps_patterns_extra/map_bool_keys_runtime
// origin: languages/go/tests/go/test_maps_patterns_extra.rs

package main
import "fmt"
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { values := map[bool]string{true: "yes", false: "no"}
__check(fmt.Sprint(values[true]), "yes")
}
