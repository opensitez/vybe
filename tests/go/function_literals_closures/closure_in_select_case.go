// vybe-test: go/function_literals_closures/closure_in_select_case
// origin: languages/go/tests/go/test_function_literals_closures.rs

package main
import "fmt"
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { ch := make(chan int, 1)
ch <- 1
select { case fn := func(v int) int { return v }(<-ch): __check(fmt.Sprint(fn), "1") } }
