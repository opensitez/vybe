// vybe-test: go/regexp_advanced_runtime/regexp_literal_prefix_metachar_not_literal
// origin: languages/go/tests/go/test_regexp_advanced_runtime.rs

package main
import "fmt"
import "regexp"
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { re := regexp.MustCompile(`a.b`)
_, lit := re.LiteralPrefix()
__check(fmt.Sprint(lit), "false") }
