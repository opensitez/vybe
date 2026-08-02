// vybe-test: go/defer_lifo_extended/defer_named_return_two_values_both_set
// origin: languages/go/tests/go/test_defer_lifo_extended.rs

package main
import "fmt"
func work() (a int, b int) { defer func() { a = 1
b = 2 }()
return 9, 8 }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { x, y := work()
__check(fmt.Sprint(x), "1")
__check(fmt.Sprint(y), "2") }
