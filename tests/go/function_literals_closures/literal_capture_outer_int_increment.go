// vybe-test: go/function_literals_closures/literal_capture_outer_int_increment
// origin: languages/go/tests/go/test_function_literals_closures.rs

package main
import "fmt"
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { n := 0
inc := func() { n++ }
inc()
inc()
__check(fmt.Sprint(n), "2") }
