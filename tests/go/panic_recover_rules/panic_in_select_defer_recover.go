// vybe-test: go/panic_recover_rules/panic_in_select_defer_recover
// origin: languages/go/tests/go/test_panic_recover_rules.rs

package main
import "fmt"
func run() { defer func() { __check(fmt.Sprint(recover()), "sel") }()
ch := make(chan int, 1)
ch <- 1
select { case <-ch: panic("sel") } }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { run() }
