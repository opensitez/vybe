// vybe-test: go/sort_stable_search/sort_slice_struct_by_field
// origin: languages/go/tests/go/test_sort_stable_search.rs

package main
import "fmt"
import "sort"
type pair struct { k, v int }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { s := []pair{{3, 30}, {1, 10}, {2, 20}}
sort.Slice(s, func(i, j int) bool { return s[i].k < s[j].k })
__check(fmt.Sprint(s[0].v), "10")
__check(fmt.Sprint(s[2].v), "30") }
