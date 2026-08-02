// vybe-test: go/defer_lifo_extended/defer_lifo_preserves_registration_order_on_panic
// origin: languages/go/tests/go/test_defer_lifo_extended.rs

package main
import "fmt"
func run() { defer __check(fmt.Sprint(1), "2")
defer __check(fmt.Sprint(2), "1")
defer func() { recover() }()
panic("x") }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { run() }
