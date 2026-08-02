// vybe-test: go/select_patterns_advanced/select_closed_after_draining_buffered_value
// origin: languages/go/tests/go/test_select_patterns_advanced.rs

package main
import "fmt"
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { ch := make(chan int, 2)
ch <- 10
ch <- 20
close(ch)
select { case v := <-ch: __check(fmt.Sprint(v), "10") }
select { case v := <-ch: __check(fmt.Sprint(v), "20") }
select { case v, ok := <-ch: __check(fmt.Sprint(v), "0")
__check(fmt.Sprint(ok), "false") } }
