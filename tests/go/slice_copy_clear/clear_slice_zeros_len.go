// vybe-test: go/slice_copy_clear/clear_slice_zeros_len
// origin: languages/go/tests/go/test_slice_copy_clear.rs

package main
import "fmt"
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { s := []int{1,2,3}
clear(s)
__check(fmt.Sprint(len(s)), "3") }
