// vybe-test: go/slices_delete_insert/slices_grow
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

func main() { s := []int{1}
t := slices.Grow(s, 2)
__check(fmt.Sprint(len(t)) + " " + fmt.Sprint(cap(t) >= 3), "1 true") }
