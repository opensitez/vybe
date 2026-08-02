// vybe-test: go/blank_identifier_extended/blank_multi_assign_keep_first
// origin: languages/go/tests/go/test_blank_identifier_extended.rs

package main
import "fmt"
func pair() (int, string) { return 7, "go" }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { a, _ := pair()
__check(fmt.Sprint(a), "7") }
