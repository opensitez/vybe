// vybe-test: go/interface_nil_comparable/generic_comparable_array_equality
// origin: languages/go/tests/go/test_interface_nil_comparable.rs

package main
import "fmt"
func equalArray(left [2]int, right [2]int) bool { return left == right }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { __check(fmt.Sprint(equalArray([2]int{1, 2}, [2]int{1, 2})), "true")
__check(fmt.Sprint(equalArray([2]int{1, 2}, [2]int{2, 1})), "false") }
