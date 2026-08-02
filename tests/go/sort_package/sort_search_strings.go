// vybe-test: go/sort_package/sort_search_strings
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

func main() { s := []string{"a","c","e"}
__check(fmt.Sprint(sort.SearchStrings(s, "c")), "1") }
