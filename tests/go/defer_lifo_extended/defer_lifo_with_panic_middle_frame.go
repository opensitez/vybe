// vybe-test: go/defer_lifo_extended/defer_lifo_with_panic_middle_frame
// origin: languages/go/tests/go/test_defer_lifo_extended.rs

package main
import "fmt"
func run() { defer __check(fmt.Sprint("c"), "a")
defer __check(fmt.Sprint("b"), "c")
defer func() { recover() }()
__check(fmt.Sprint("a"), "b")
panic("x") }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { run() }
