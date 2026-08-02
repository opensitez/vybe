// vybe-test: go/regexp_package/regexp_replace_all
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

func main() { re := regexp.MustCompile(`a+`)
__check(fmt.Sprint(re.ReplaceAllString("baaac", "X")), "bXc") }
