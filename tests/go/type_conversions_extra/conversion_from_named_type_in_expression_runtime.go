// vybe-test: go/type_conversions_extra/conversion_from_named_type_in_expression_runtime
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

func main() { var value score = 24
__check(fmt.Sprint(int(value) + 1), "25")
}
