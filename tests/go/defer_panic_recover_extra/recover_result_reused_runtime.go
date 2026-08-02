// vybe-test: go/defer_panic_recover_extra/recover_result_reused_runtime
// origin: languages/go/tests/go/test_defer_panic_recover_extra.rs

package main
import "fmt"
func run() { defer func() { value := recover()
__check(fmt.Sprint(value), "3")
__check(fmt.Sprint(value != nil), "true") }()
panic(3) }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { run() }
