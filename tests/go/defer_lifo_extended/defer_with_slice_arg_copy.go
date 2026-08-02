// vybe-test: go/defer_lifo_extended/defer_with_slice_arg_copy
// origin: languages/go/tests/go/test_defer_lifo_extended.rs

package main
import "fmt"
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { s := []int{1, 2}
defer __check(fmt.Sprint(len(s)), "2")
s = append(s, 3)
}
