// vybe-test: go/type_conversions_extra/conversion_inside_function_call_runtime
// origin: languages/go/tests/go/test_type_conversions_extra.rs

package main
import "fmt"
func show(v float64) { __check(fmt.Sprint(v), "23") }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { show(float64(23))
}
