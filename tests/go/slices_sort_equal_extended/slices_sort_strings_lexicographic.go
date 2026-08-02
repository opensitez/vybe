// vybe-test: go/slices_sort_equal_extended/slices_sort_strings_lexicographic
// origin: languages/go/tests/go/test_slices_sort_equal_extended.rs

package main
import "fmt"
import "slices"
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { s := []string{"cherry", "apple", "banana"}
slices.Sort(s)
__check(fmt.Sprint(s[0]), "apple")
__check(fmt.Sprint(s[2]), "cherry") }
