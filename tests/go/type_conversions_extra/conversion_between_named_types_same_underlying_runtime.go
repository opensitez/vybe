// vybe-test: go/type_conversions_extra/conversion_between_named_types_same_underlying_runtime
// origin: languages/go/tests/go/test_type_conversions_extra.rs

package main
import "fmt"
type first int
type second int
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { var value first = 19
__check(fmt.Sprint(second(value)), "19")
}
