// vybe-test: go/defer_lifo_extended/defer_captures_outer_var_by_reference
// origin: languages/go/tests/go/test_defer_lifo_extended.rs

package main
import "fmt"
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { n := 0
defer func() { __check(fmt.Sprint(n), "7") }()
n = 7
}
