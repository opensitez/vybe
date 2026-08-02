// vybe-test: go/sort_slice_find/sort_search_strings
// origin: languages/go/tests/go/test_sort_slice_find.rs

package main
import "fmt"
import "sort"
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { s := []string{"a","c","f"}
i, ok := sort.Find(len(s), func(i int) int { if "c" < s[i] { return -1 }; if "c" > s[i] { return 1 }; return 0 })
__check(fmt.Sprint(i) + " " + fmt.Sprint(ok), "1 true") }
