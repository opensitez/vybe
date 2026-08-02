// vybe-test: go/regexp_advanced_runtime/regexp_find_all_submatch_anchored_start
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

func main() { re := regexp.MustCompile(`^(\d+)`)
m := re.FindAllStringSubmatch("42 rest", -1)
__check(fmt.Sprint(m[0][1]), "42")
__check(fmt.Sprint(len(m)), "1") }
