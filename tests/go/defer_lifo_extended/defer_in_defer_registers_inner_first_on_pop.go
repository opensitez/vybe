// vybe-test: go/defer_lifo_extended/defer_in_defer_registers_inner_first_on_pop
// origin: languages/go/tests/go/test_defer_lifo_extended.rs

package main
import "fmt"
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { defer func() { defer __check(fmt.Sprint(2), "1")
__check(fmt.Sprint(1), "2") }()
}
