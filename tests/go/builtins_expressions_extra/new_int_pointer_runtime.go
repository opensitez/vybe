// vybe-test: go/builtins_expressions_extra/new_int_pointer_runtime
// origin: languages/go/tests/go/test_builtins_expressions_extra.rs

package main
import "fmt"
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { value := new(int)
*value = 11
__check(fmt.Sprint(*value), "11")
}
