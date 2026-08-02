// vybe-test: go/sort_stable_search/sort_slice_by_string_length
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

func main() { s := []string{"go", "vybe", "a", "lang"}
sort.Slice(s, func(i, j int) bool { return len(s[i]) < len(s[j]) })
__check(fmt.Sprint(s[0]), "a")
__check(fmt.Sprint(s[3]), "vybe") }
