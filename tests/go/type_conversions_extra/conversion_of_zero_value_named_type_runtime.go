// vybe-test: go/type_conversions_extra/conversion_of_zero_value_named_type_runtime
// origin: languages/go/tests/go/test_type_conversions_extra.rs

package main
import "fmt"
type score int
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { var value score
__check(fmt.Sprint(int(value)), "0")
}
