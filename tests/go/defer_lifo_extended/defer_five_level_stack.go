// vybe-test: go/defer_lifo_extended/defer_five_level_stack
// origin: languages/go/tests/go/test_defer_lifo_extended.rs

package main
import "fmt"
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { defer __check(fmt.Sprint("a"), "a")
defer __check(fmt.Sprint("b"), "b")
defer __check(fmt.Sprint("c"), "c")
defer __check(fmt.Sprint("d"), "d")
defer __check(fmt.Sprint("e"), "e")
}
