// vybe-test: go/type_conversions_extra/conversion_in_short_decl_runtime
// origin: languages/go/tests/go/test_type_conversions_extra.rs

package main
import "fmt"
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { value := int(18.2)
__check(fmt.Sprint(value), "18")
}
