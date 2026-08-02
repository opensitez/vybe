// vybe-test: go/sort_stable_search/sort_ints_mixed_positive_negative
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

func main() { a := []int{10, -10, 0, 5, -5}
sort.Ints(a)
__check(fmt.Sprint(a[0]), "-10")
__check(fmt.Sprint(a[2]), "0")
__check(fmt.Sprint(a[4]), "10") }
