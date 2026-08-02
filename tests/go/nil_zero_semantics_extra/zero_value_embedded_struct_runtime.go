// vybe-test: go/nil_zero_semantics_extra/zero_value_embedded_struct_runtime
// origin: languages/go/tests/go/test_nil_zero_semantics_extra.rs

package main
import "fmt"
type inner struct { count int }
type outer struct { inner }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { var value outer
__check(fmt.Sprint(value.count), "0")
}
