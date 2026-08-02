// vybe-test: go/sort_stable_search/stable_sort_three_equal_groups
// origin: languages/go/tests/go/test_sort_stable_search.rs

package main
import "fmt"
import "sort"
type rec struct { g, seq int }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { s := []rec{{0, 0}, {1, 0}, {0, 1}, {1, 1}, {0, 2}}
sort.SliceStable(s, func(i, j int) bool { return s[i].g < s[j].g })
__check(fmt.Sprint(s[0].seq), "0")
__check(fmt.Sprint(s[1].seq), "1")
__check(fmt.Sprint(s[2].seq), "2") }
