// vybe-test: go/fmt_sprintf_verbs/sprintf_binary
// origin: languages/go/tests/go/test_fmt_sprintf_verbs.rs

package main
import "fmt"
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { __check(fmt.Sprint(fmt.Sprintf("%b", 5)), "101") }
