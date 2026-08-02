// vybe-test: go/regexp_advanced_runtime/regexp_find_all_submatch_word_boundary
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

func main() { re := regexp.MustCompile(`\b(\w{2})\b`)
m := re.FindAllStringSubmatch("go is ok", -1)
__check(fmt.Sprint(len(m)), "3")
__check(fmt.Sprint(m[0][1]), "go") }
