// vybe-test: go/sort_stable_search/stable_sort_already_sorted_stable
// origin: languages/go/tests/go/test_sort_stable_search.rs

package main
import "fmt"
import "sort"
type kv struct { k, ord int }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { s := []kv{{1, 0}, {2, 1}, {3, 2}}
sort.SliceStable(s, func(i, j int) bool { return s[i].k < s[j].k })
__check(fmt.Sprint(s[0].ord), "0")
__check(fmt.Sprint(s[2].ord), "2") }
