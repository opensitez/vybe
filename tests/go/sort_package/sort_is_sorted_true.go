// vybe-test: go/sort_package/sort_is_sorted_true
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

func main() { __check(fmt.Sprint(sort.IntsAreSorted([]int{1,2,3})), "true") }
