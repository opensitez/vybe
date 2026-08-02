// vybe-test: go/function_literals_closures/closure_capture_pointer
// origin: languages/go/tests/go/test_function_literals_closures.rs

package main
import "fmt"
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { n := 5
ptr := &n
bump := func() { *ptr = *ptr + 1 }
bump()
__check(fmt.Sprint(n), "6") }
