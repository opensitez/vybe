// vybe-test: go/nil_zero_semantics_extra/zero_value_named_type_runtime
// origin: languages/go/tests/go/test_nil_zero_semantics_extra.rs

package main
import "fmt"
type score int
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { var s score
__check(fmt.Sprint(s), "0")
}
