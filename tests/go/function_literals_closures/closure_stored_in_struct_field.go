// vybe-test: go/function_literals_closures/closure_stored_in_struct_field
// origin: languages/go/tests/go/test_function_literals_closures.rs

package main
import "fmt"
type holder struct { fn func() int }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { h := holder{fn: func() int { return 42 }}
__check(fmt.Sprint(h.fn()), "42") }
