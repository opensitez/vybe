// vybe-test: go/type_conversions_extra/array_value_conversion_runtime
// origin: languages/go/tests/go/test_type_conversions_extra.rs

package main
import "fmt"
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { values := [2]int{1, 14}
__check(fmt.Sprint(float64(values[1])), "14")
}
