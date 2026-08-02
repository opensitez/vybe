// vybe-test: go/defer_lifo_extended/defer_named_return_float64_scaled
// origin: languages/go/tests/go/test_defer_lifo_extended.rs

package main
import "fmt"
func work() (f float64) { defer func() { f = f * 2 }()
return 3.5 }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { __check(fmt.Sprint(f == 7.0), "true") }
