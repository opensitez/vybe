// vybe-test: go/defer_panic_recover_extra/panic_after_defer_runtime
// origin: languages/go/tests/go/test_defer_panic_recover_extra.rs

package main
import "fmt"
func run() { defer __check(fmt.Sprint("cleanup"), "cleanup")
defer func() { recover() }()
panic("stop") }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { run() }
