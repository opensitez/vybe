// vybe-test: go/defer_lifo_extended/defer_arg_evaluated_at_register_time
// origin: languages/go/tests/go/test_defer_lifo_extended.rs

package main
import "fmt"
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { n := 1
defer __check(fmt.Sprint(n), "1")
n = 9
}
