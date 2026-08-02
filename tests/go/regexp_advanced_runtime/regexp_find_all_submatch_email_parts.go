// vybe-test: go/regexp_advanced_runtime/regexp_find_all_submatch_email_parts
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

func main() { re := regexp.MustCompile(`([\w.]+)@([\w.]+)`)
m := re.FindAllStringSubmatch("a@b.com c@d.org", -1)
__check(fmt.Sprint(m[0][1]), "a")
__check(fmt.Sprint(m[1][2]), "d.org") }
