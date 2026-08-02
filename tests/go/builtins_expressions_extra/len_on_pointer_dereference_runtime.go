// vybe-test: go/builtins_expressions_extra/len_on_pointer_dereference_runtime
// origin: languages/go/tests/go/test_builtins_expressions_extra.rs

package main
import "fmt"
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { values := &[3]int{1, 2, 3}
__check(fmt.Sprint(len(*values)), "3")
}
