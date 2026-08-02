// vybe-test: go/sort_package/sort_slice_unstable_desc
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

func main() { a := []int{1,3,2}
sort.Slice(a, func(i,j int) bool { return a[i] > a[j] })
__check(fmt.Sprint(a[0]), "3")
__check(fmt.Sprint(a[2]), "1") }
