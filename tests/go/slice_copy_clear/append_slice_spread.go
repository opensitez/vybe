// vybe-test: go/slice_copy_clear/append_slice_spread
// origin: languages/go/tests/go/test_slice_copy_clear.rs

package main
import "fmt"
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { base := []int{1}
more := []int{2,3}
s := append(base, more...)
__check(fmt.Sprint(len(s)), "3")
__check(fmt.Sprint(s[2]), "3") }
