// vybe-test: go/select_patterns_advanced/select_closed_buffered_drains_value_then_zero
// origin: languages/go/tests/go/test_select_patterns_advanced.rs

package main
import "fmt"
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { ch := make(chan int, 1)
ch <- 42
close(ch)
select { case v := <-ch: __check(fmt.Sprint(v), "42") }
select { case v, ok := <-ch: __check(fmt.Sprint(v), "0")
__check(fmt.Sprint(ok), "false") } }
