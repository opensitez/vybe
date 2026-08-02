// vybe-test: go/nil_zero_semantics_extra/nil_channel_compare_runtime
// origin: languages/go/tests/go/test_nil_zero_semantics_extra.rs

package main
import "fmt"
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { var ch chan int
__check(fmt.Sprint(ch == nil), "true")
}
