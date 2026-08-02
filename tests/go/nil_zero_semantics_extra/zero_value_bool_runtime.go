// vybe-test: go/nil_zero_semantics_extra/zero_value_bool_runtime
// origin: languages/go/tests/go/test_nil_zero_semantics_extra.rs

package main
import "fmt"
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { var flag bool
__check(fmt.Sprint(flag), "false")
}
