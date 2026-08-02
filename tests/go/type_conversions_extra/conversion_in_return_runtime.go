// vybe-test: go/type_conversions_extra/conversion_in_return_runtime
// origin: languages/go/tests/go/test_type_conversions_extra.rs

package main
import "fmt"
func cast(v int) float64 { return float64(v) }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { __check(fmt.Sprint(cast(5)), "5")
}
