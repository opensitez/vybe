// vybe-test: go/slices_delete_insert/slices_delete_middle
// origin: languages/go/tests/go/test_slices_delete_insert.rs

package main
import "fmt"
import "slices"
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { s := []int{1,2,3,4}
t := slices.Delete(s, 1, 3)
__check(fmt.Sprint(t), "[1 4]") }
