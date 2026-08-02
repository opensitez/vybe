// vybe-test: go/lang_expressions/make_chan_buffered
// origin: languages/go/tests/go/test_lang_expressions.rs

package main
import "fmt"
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { ch := make(chan int, 2)
__check(fmt.Sprint(cap(ch)), "2") }
