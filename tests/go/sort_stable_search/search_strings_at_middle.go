// vybe-test: go/sort_stable_search/search_strings_at_middle
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

func main() { s := []string{"alpha", "beta", "gamma"}
__check(fmt.Sprint(sort.SearchStrings(s, "beta")), "1") }
