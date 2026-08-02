// vybe-test: go/regexp_advanced_runtime/regexp_find_all_submatch_non_capturing
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

func main() { re := regexp.MustCompile(`(?:ab)+`)
m := re.FindAllStringSubmatch("abab", -1)
__check(fmt.Sprint(len(m[0])), "1")
__check(fmt.Sprint(m[0][0]), "abab") }
