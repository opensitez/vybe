// vybe-test: go/interface_nil_comparable/make_channel_not_equal_nil
// origin: languages/go/tests/go/test_interface_nil_comparable.rs

package main
import "fmt"
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { ch := make(chan int)
__check(fmt.Sprint(ch == nil), "false") }
