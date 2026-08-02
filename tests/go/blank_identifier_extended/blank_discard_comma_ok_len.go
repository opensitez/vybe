// vybe-test: go/blank_identifier_extended/blank_discard_comma_ok_len
// origin: languages/go/tests/go/test_blank_identifier_extended.rs

package main
import "fmt"
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { s := []int{1, 2, 3}
_ = len(s)
__check(fmt.Sprint(s[0]), "1") }
