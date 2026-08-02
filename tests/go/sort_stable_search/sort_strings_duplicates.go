// vybe-test: go/sort_stable_search/sort_strings_duplicates
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

func main() { s := []string{"z", "a", "z", "m"}
sort.Strings(s)
__check(fmt.Sprint(s[0]), "a")
__check(fmt.Sprint(s[3]), "z") }
