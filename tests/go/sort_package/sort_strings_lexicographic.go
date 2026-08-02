// vybe-test: go/sort_package/sort_strings_lexicographic
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

func main() { s := []string{"b","a","c"}
sort.Strings(s)
__check(fmt.Sprint(s[0]), "a")
__check(fmt.Sprint(s[2]), "c") }
