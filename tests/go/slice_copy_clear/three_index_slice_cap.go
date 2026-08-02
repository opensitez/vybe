// vybe-test: go/slice_copy_clear/three_index_slice_cap
// origin: languages/go/tests/go/test_slice_copy_clear.rs

package main
import "fmt"
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { a := []int{0,1,2,3,4}
b := a[1:3:4]
__check(fmt.Sprint(len(b)), "2")
__check(fmt.Sprint(cap(b)), "3")
__check(fmt.Sprint(b[1]), "2") }
