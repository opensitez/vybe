// vybe-test: go/lang_declarations_types/channel_eq_nil
// origin: languages/go/tests/go/test_lang_declarations_types.rs

package main
import "fmt"
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { var ch chan int
__check(fmt.Sprint(ch == nil), "true") }
