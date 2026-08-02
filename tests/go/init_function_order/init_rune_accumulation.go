// vybe-test: go/init_function_order/init_rune_accumulation
// origin: languages/go/tests/go/test_init_function_order.rs

package main
import "fmt"
var ch rune
func init() { ch = 'A' }
func init() { ch = ch + 1 }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { __check(fmt.Sprint(int(ch)), "66") }
