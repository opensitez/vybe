// vybe-test: go/defer_lifo_extended/defer_two_prints_reverse_order
// origin: languages/go/tests/go/test_defer_lifo_extended.rs

package main
import "fmt"
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { defer __check(fmt.Sprint("first"), "first")
defer __check(fmt.Sprint("second"), "second")
}
