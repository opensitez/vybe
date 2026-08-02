// vybe-test: go/sort_stable_search/reverse_string_slice
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

func main() { s := sort.StringSlice{"a", "b", "c"}
sort.Sort(sort.Reverse(s))
__check(fmt.Sprint(s[0]), "c")
__check(fmt.Sprint(s[2]), "a") }
