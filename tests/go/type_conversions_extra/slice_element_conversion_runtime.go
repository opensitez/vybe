// vybe-test: go/type_conversions_extra/slice_element_conversion_runtime
// origin: languages/go/tests/go/test_type_conversions_extra.rs

package main
import "fmt"
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { values := []int{4, 6}
__check(fmt.Sprint(float64(values[1])), "6")
}
