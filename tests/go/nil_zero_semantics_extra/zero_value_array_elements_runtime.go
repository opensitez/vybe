// vybe-test: go/nil_zero_semantics_extra/zero_value_array_elements_runtime
// origin: languages/go/tests/go/test_nil_zero_semantics_extra.rs

package main
import "fmt"
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { var values [2]int
__check(fmt.Sprint(values[0]), "0")
__check(fmt.Sprint(values[1]), "0")
}
