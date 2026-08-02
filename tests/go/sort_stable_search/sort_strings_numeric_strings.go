// vybe-test: go/sort_stable_search/sort_strings_numeric_strings
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

func main() { s := []string{"10", "2", "1"}
sort.Strings(s)
__check(fmt.Sprint(s[0]), "1")
__check(fmt.Sprint(s[2]), "2") }
