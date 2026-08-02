// vybe-test: go/type_conversions_extra/int_literal_named_type_runtime
// origin: languages/go/tests/go/test_type_conversions_extra.rs

package main
import "fmt"
type level int
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { value := level(16)
__check(fmt.Sprint(value), "16")
}
