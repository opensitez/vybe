// vybe-test: go/nil_zero_semantics_extra/nil_slice_len_runtime
// origin: languages/go/tests/go/test_nil_zero_semantics_extra.rs

package main
import "fmt"
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { var values []int
__check(fmt.Sprint(values == nil), "true")
__check(fmt.Sprint(len(values)), "0")
}
