// vybe-test: go/defer_lifo_extended/defer_chain_three_closures
// origin: languages/go/tests/go/test_defer_lifo_extended.rs

package main
import "fmt"
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { defer func() { __check(fmt.Sprint(3), "3") }()
defer func() { __check(fmt.Sprint(2), "2") }()
defer func() { __check(fmt.Sprint(1), "1") }()
}
