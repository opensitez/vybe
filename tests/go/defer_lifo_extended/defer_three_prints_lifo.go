// vybe-test: go/defer_lifo_extended/defer_three_prints_lifo
// origin: languages/go/tests/go/test_defer_lifo_extended.rs

package main
import "fmt"
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { defer __check(fmt.Sprint(1), "1")
defer __check(fmt.Sprint(2), "2")
defer __check(fmt.Sprint(3), "3")
}
