// vybe-test: go/nil_zero_semantics_extra/zero_value_array_field_runtime
// origin: languages/go/tests/go/test_nil_zero_semantics_extra.rs

package main
import "fmt"
type bag struct { values [2]int }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { var b bag
__check(fmt.Sprint(b.values[1]), "0")
}
