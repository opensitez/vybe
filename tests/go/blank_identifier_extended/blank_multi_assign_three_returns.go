// vybe-test: go/blank_identifier_extended/blank_multi_assign_three_returns
// origin: languages/go/tests/go/test_blank_identifier_extended.rs

package main
import "fmt"
func triple() (int, int, int) { return 1, 2, 3 }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { _, mid, _ := triple()
__check(fmt.Sprint(mid), "2") }
