// vybe-test: go/nil_zero_semantics_extra/zero_value_channel_in_struct_runtime
// origin: languages/go/tests/go/test_nil_zero_semantics_extra.rs

package main
import "fmt"
type holder struct { ch chan int }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { var h holder
__check(fmt.Sprint(h.ch == nil), "true")
}
