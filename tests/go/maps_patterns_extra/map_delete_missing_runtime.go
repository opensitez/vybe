// vybe-test: go/maps_patterns_extra/map_delete_missing_runtime
// origin: languages/go/tests/go/test_maps_patterns_extra.rs

package main
import "fmt"
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { values := map[string]int{"a": 1}
delete(values, "missing")
__check(fmt.Sprint(len(values)), "1")
}
