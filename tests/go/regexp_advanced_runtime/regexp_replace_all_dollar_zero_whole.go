// vybe-test: go/regexp_advanced_runtime/regexp_replace_all_dollar_zero_whole
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

func main() { re := regexp.MustCompile(`(\d+)`)
__check(fmt.Sprint(re.ReplaceAllString("x1y2", "[$0]")), "x[1]y[2]") }
