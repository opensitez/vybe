// vybe-test: go/type_conversions_extra/alias_type_conversion_runtime
// origin: languages/go/tests/go/test_type_conversions_extra.rs

package main
import "fmt"
type count = int
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { value := count(9)
__check(fmt.Sprint(value), "9")
}
