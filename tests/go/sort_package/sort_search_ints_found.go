// vybe-test: go/sort_package/sort_search_ints_found
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

func main() { a := []int{1,3,5}
__check(fmt.Sprint(sort.SearchInts(a, 3)), "1") }
