// vybe-test: go/regexp_advanced_runtime/regexp_literal_prefix_fixed_text
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

func main() { re := regexp.MustCompile(`prefix(\d+)`)
p, lit := re.LiteralPrefix()
__check(fmt.Sprint(p), "prefix")
__check(fmt.Sprint(lit), "false") }
