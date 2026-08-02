// vybe-test: go/sort_stable_search/stable_sort_preserves_equal_key_order_first
// origin: languages/go/tests/go/test_sort_stable_search.rs

package main
import "fmt"
import "sort"
type tagged struct { key, ord int }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { s := []tagged{{1, 0}, {2, 0}, {1, 1}, {1, 2}}
sort.SliceStable(s, func(i, j int) bool { return s[i].key < s[j].key })
__check(fmt.Sprint(s[1].ord), "0")
__check(fmt.Sprint(s[2].ord), "1")
__check(fmt.Sprint(s[3].ord), "2") }
