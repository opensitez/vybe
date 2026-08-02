// vybe-test: go/maps_patterns_extra/nil_map_lookup_len_runtime
// origin: languages/go/tests/go/test_maps_patterns_extra.rs

package main
import "fmt"
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { var values map[string]int
__check(fmt.Sprint(values["a"]), "0")
__check(fmt.Sprint(len(values)), "0")
}
