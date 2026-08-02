// vybe-test: go/regexp_advanced_runtime/regexp_find_all_submatch_word_triple
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

func main() { re := regexp.MustCompile(`(\w)(\w)(\w)`)
m := re.FindAllStringSubmatch("goo!", -1)
__check(fmt.Sprint(len(m[0])), "4")
__check(fmt.Sprint(m[0][3]), "o") }
