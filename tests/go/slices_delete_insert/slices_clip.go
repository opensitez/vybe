// vybe-test: go/slices_delete_insert/slices_clip
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

func main() { s := make([]int, 3, 10)
t := slices.Clip(s)
__check(fmt.Sprint(len(t)) + " " + fmt.Sprint(cap(t)), "3 3") }
