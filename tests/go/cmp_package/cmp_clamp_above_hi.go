// vybe-test: go/cmp_package/cmp_clamp_above_hi
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

func main() { __check(fmt.Sprint(cmp.Clamp(99, 0, 9)), "9") }
