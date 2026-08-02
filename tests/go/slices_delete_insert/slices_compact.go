// vybe-test: go/slices_delete_insert/slices_compact
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

func main() { s := []int{1,0,2,0,3}
t := slices.Compact(s)
__check(fmt.Sprint(t), "[1 0 2 0 3]") }
