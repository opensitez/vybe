// vybe-test: go/regexp_advanced_runtime/regexp_quote_meta_dot_star
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

func main() { __check(fmt.Sprint(regexp.QuoteMeta("a.b*")), "a\\.b\\*") }
