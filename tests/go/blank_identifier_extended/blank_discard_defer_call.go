// vybe-test: go/blank_identifier_extended/blank_discard_defer_call
// origin: languages/go/tests/go/test_blank_identifier_extended.rs

package main
import "fmt"
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { x := 1
defer func() { _ = x }()
__check(fmt.Sprint(x), "1") }
