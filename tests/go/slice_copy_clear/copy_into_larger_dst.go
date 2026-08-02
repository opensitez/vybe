// vybe-test: go/slice_copy_clear/copy_into_larger_dst
// origin: languages/go/tests/go/test_slice_copy_clear.rs

package main
import "fmt"
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { dst := make([]int, 5)
src := []int{7,8}
n := copy(dst, src)
__check(fmt.Sprint(n), "2")
__check(fmt.Sprint(dst[0]), "7")
__check(fmt.Sprint(dst[4]), "0") }
