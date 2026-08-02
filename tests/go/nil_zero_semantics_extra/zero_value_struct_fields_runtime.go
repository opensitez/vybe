// vybe-test: go/nil_zero_semantics_extra/zero_value_struct_fields_runtime
// origin: languages/go/tests/go/test_nil_zero_semantics_extra.rs

package main
import "fmt"
type counter struct { n int
ok bool }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { var c counter
__check(fmt.Sprint(c.n), "0")
__check(fmt.Sprint(c.ok), "false")
}
