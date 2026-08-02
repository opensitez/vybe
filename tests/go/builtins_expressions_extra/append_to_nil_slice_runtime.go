// vybe-test: go/builtins_expressions_extra/append_to_nil_slice_runtime
// origin: languages/go/tests/go/test_builtins_expressions_extra.rs

package main
import "fmt"
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { var values []int
values = append(values, 9)
__check(fmt.Sprint(len(values)), "1")
__check(fmt.Sprint(values[0]), "9")
}
