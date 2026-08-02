// vybe-test: go/sort_stable_search/sort_ints_all_equal
// origin: languages/go/tests/go/test_sort_stable_search.rs

package main
import "fmt"
import "sort"
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { a := []int{5, 5, 5}
sort.Ints(a)
__check(fmt.Sprint(a[0]), "5")
__check(fmt.Sprint(len(a)), "3") }
