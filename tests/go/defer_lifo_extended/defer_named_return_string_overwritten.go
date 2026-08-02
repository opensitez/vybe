// vybe-test: go/defer_lifo_extended/defer_named_return_string_overwritten
// origin: languages/go/tests/go/test_defer_lifo_extended.rs

package main
import "fmt"
func work() (s string) { defer func() { s = "bye" }()
return "hi" }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { __check(fmt.Sprint(work()), "bye") }
