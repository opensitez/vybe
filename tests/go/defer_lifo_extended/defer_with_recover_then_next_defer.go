// vybe-test: go/defer_lifo_extended/defer_with_recover_then_next_defer
// origin: languages/go/tests/go/test_defer_lifo_extended.rs

package main
import "fmt"
func run() { defer __check(fmt.Sprint("last"), "last")
defer func() { recover() }()
defer __check(fmt.Sprint("mid"), "mid")
panic("p") }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { run() }
