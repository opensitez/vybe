// vybe-test: go/maps_patterns_extra/map_make_set_len_runtime
// origin: languages/go/tests/go/test_maps_patterns_extra.rs

package main
import "fmt"
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { values := make(map[string]int)
values["a"] = 3
__check(fmt.Sprint(len(values)), "1")
}
