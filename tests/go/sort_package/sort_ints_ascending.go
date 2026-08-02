// vybe-test: go/sort_package/sort_ints_ascending
// origin: languages/go/tests/go/test_sort_package.rs

package main
import "fmt"
import "sort"
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { a := []int{3,1,2}
sort.Ints(a)
__check(fmt.Sprint(a[0]), "1")
__check(fmt.Sprint(a[2]), "3") }
