// vybe-test: go/nil_zero_semantics_extra/nil_pointer_compare_runtime
// origin: languages/go/tests/go/test_nil_zero_semantics_extra.rs

package main
import "fmt"
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { var value *int
__check(fmt.Sprint(value == nil), "true")
}
