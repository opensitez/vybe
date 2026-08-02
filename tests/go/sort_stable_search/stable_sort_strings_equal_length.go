// vybe-test: go/sort_stable_search/stable_sort_strings_equal_length
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

func main() { s := []string{"bb", "aa", "cc", "dd"}
sort.SliceStable(s, func(i, j int) bool { return len(s[i]) < len(s[j]) })
__check(fmt.Sprint(s[0]), "aa")
__check(fmt.Sprint(s[3]), "dd") }
