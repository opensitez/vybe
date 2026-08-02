// vybe-test: go/cmp_package/cmp_compare_int_less
// origin: languages/go/tests/go/test_cmp_package.rs

package main
import "fmt"
import "cmp"
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { __check(fmt.Sprint(cmp.Compare(1, 2)), "-1") }
