// vybe-test: go/sort_stable_search/sort_strings_empty_element
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

func main() { s := []string{"b", "", "a"}
sort.Strings(s)
__check(fmt.Sprint(s[0]), "")
__check(fmt.Sprint(s[2]), "b") }
