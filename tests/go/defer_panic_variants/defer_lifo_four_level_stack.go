// vybe-test: go/defer_panic_variants/defer_lifo_four_level_stack
// origin: languages/go/tests/go/test_defer_panic_variants.rs

package main
import "fmt"
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { defer __check(fmt.Sprint("a"), "d")
defer __check(fmt.Sprint("b"), "c")
defer __check(fmt.Sprint("c"), "b")
defer __check(fmt.Sprint("d"), "a")
}
