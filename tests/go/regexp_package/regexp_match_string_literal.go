// vybe-test: go/regexp_package/regexp_match_string_literal
// origin: languages/go/tests/go/test_regexp_package.rs

package main
import "fmt"
import "regexp"
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { __check(fmt.Sprint(regexp.MatchString("^go", "gopher")), "true") }
