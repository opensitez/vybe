// vybe-test: go/defer_lifo_extended/defer_in_anonymous_func_separate_stack
// origin: languages/go/tests/go/test_defer_lifo_extended.rs

package main
import "fmt"
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { defer __check(fmt.Sprint("main"), "anon")
func() { defer __check(fmt.Sprint("anon"), "main")
}()
}
