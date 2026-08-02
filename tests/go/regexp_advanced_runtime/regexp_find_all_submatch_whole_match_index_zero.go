// vybe-test: go/regexp_advanced_runtime/regexp_find_all_submatch_whole_match_index_zero
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
m := re.FindAllStringSubmatch("n7", -1)
__check(fmt.Sprint(m[0][0]), "7") }
