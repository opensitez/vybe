// vybe-test: go/regexp_package/regexp_find_first_submatch
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

func main() { re := regexp.MustCompile(`(\d+)`)
m := re.FindStringSubmatch("id:42")
__check(fmt.Sprint(m[1]), "42") }
