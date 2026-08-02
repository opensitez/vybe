// vybe-test: go/type_conversions_extra/float32_to_float64_runtime
// origin: languages/go/tests/go/test_type_conversions_extra.rs

package main
import "fmt"
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { var value float32 = 15
__check(fmt.Sprint(float64(value)), "15")
}
