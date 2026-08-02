// vybe-test: go/maps_patterns_extra/map_alias_mutation_visible_runtime
// origin: languages/go/tests/go/test_maps_patterns_extra.rs

package main
import "fmt"
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { left := map[string]int{"a": 1}
right := left
right["a"] = 9
__check(fmt.Sprint(left["a"]), "9")
}
