// vybe-test: go/sort_stable_search/search_strings_before_all
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

func main() { s := []string{"m", "n", "o"}
__check(fmt.Sprint(sort.SearchStrings(s, "a")), "0") }
