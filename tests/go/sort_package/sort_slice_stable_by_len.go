// vybe-test: go/sort_package/sort_slice_stable_by_len
// origin: languages/go/tests/go/test_sort_package.rs

package main
import "fmt"
import "sort"
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { items := []string{"aaa","b","cc"}
sort.SliceStable(items, func(i,j int) bool { return len(items[i]) < len(items[j]) })
__check(fmt.Sprint(items[0]), "b")
__check(fmt.Sprint(items[2]), "aaa") }
