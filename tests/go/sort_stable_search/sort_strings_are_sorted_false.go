// vybe-test: go/sort_stable_search/sort_strings_are_sorted_false
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

func main() { __check(fmt.Sprint(sort.StringsAreSorted([]string{"c", "a"})), "false") }
