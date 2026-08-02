// vybe-test: go/regexp_advanced_runtime/regexp_literal_prefix_no_literal
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

func main() { re := regexp.MustCompile(`^\d+`)
p, lit := re.LiteralPrefix()
__check(fmt.Sprint(p), "")
__check(fmt.Sprint(lit), "false") }
