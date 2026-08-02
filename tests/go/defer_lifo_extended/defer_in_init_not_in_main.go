// vybe-test: go/defer_lifo_extended/defer_in_init_not_in_main
// origin: languages/go/tests/go/test_defer_lifo_extended.rs

package main
import "fmt"
var x = func() int { defer __check(fmt.Sprint("init"), "init")
return 1 }()
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { __check(fmt.Sprint(x), "1") }
