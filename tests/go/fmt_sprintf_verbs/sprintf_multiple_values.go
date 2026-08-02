// vybe-test: go/fmt_sprintf_verbs/sprintf_multiple_values
// origin: languages/go/tests/go/test_fmt_sprintf_verbs.rs

package main
import "fmt"
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { __check(fmt.Sprint(fmt.Sprintf("%s-%d", "id", 3)), "id-3") }
