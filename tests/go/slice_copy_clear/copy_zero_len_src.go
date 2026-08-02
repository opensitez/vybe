// vybe-test: go/slice_copy_clear/copy_zero_len_src
// origin: languages/go/tests/go/test_slice_copy_clear.rs

package main
import "fmt"
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { dst := []int{1,2}
n := copy(dst, []int{})
__check(fmt.Sprint(n), "0")
__check(fmt.Sprint(dst[0]), "1") }
