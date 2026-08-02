// vybe-test: go/regexp_advanced_runtime/regexp_find_all_submatch_hex_color
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

func main() { re := regexp.MustCompile(`#([0-9a-f]{2})([0-9a-f]{2})([0-9a-f]{2})`)
m := re.FindAllStringSubmatch("#aabbcc", -1)
__check(fmt.Sprint(m[0][1]), "aa")
__check(fmt.Sprint(m[0][3]), "cc") }
