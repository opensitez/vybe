// vybe-test: go/slice_aliasing_extra/slice_pass_to_function_runtime
// origin: languages/go/tests/go/test_slice_aliasing_extra.rs

package main
import "fmt"
func total(values []int) int { return values[0] + values[1] }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { __check(fmt.Sprint(total([]int{4, 5})), "9")
}
