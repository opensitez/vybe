// vybe-test: go/sort_stable_search/stable_sort_by_mod_bucket_order
// origin: languages/go/tests/go/test_sort_stable_search.rs

package main
import "fmt"
import "sort"
type item struct { v, id int }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { s := []item{{3, 0}, {1, 1}, {4, 2}, {2, 3}, {5, 4}}
sort.SliceStable(s, func(i, j int) bool { return s[i].v%2 < s[j].v%2 })
__check(fmt.Sprint(s[0].id), "1")
__check(fmt.Sprint(s[1].id), "3")
__check(fmt.Sprint(s[4].id), "4") }
