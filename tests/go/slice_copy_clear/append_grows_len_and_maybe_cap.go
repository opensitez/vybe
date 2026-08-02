// vybe-test: go/slice_copy_clear/append_grows_len_and_maybe_cap
// origin: languages/go/tests/go/test_slice_copy_clear.rs

package main
import "fmt"
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { s := make([]int, 0, 2)
s = append(s, 1, 2, 3)
__check(fmt.Sprint(len(s)), "3")
__check(fmt.Sprint(cap(s) >= 3), "true") }
