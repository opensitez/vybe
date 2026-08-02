// vybe-test: go/type_conversions_extra/map_value_conversion_runtime
// origin: languages/go/tests/go/test_type_conversions_extra.rs

package main
import "fmt"
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { values := map[string]int{"a": 13}
__check(fmt.Sprint(float64(values["a"])), "13")
}
