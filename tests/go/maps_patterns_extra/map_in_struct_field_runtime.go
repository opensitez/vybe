// vybe-test: go/maps_patterns_extra/map_in_struct_field_runtime
// origin: languages/go/tests/go/test_maps_patterns_extra.rs

package main
import "fmt"
type holder struct { values map[string]int }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { value := holder{values: map[string]int{"a": 4}}
__check(fmt.Sprint(value.values["a"]), "4")
}
