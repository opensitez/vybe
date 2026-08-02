// vybe-test: go/slice_copy_clear/copy_overlapping_slices
// origin: languages/go/tests/go/test_slice_copy_clear.rs

package main
import "fmt"
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { a := []int{1,2,3,4}
n := copy(a, a[1:])
__check(fmt.Sprint(n), "3")
__check(fmt.Sprint(a[0]), "2")
__check(fmt.Sprint(a[1]), "3") }
