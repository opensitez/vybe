// vybe-test: go/function_literals_closures/two_closures_share_outer_state
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
get := func() int { return n }
inc()
inc()
__check(fmt.Sprint(get()), "2") }
