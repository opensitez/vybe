// vybe-test: go/fmt_sprintf_verbs/sprintf_pointer_nil
// origin: languages/go/tests/go/test_fmt_sprintf_verbs.rs

package main
import "fmt"
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { var p *int
__check(fmt.Sprint(fmt.Sprintf("%p", p)), "0x0") }
