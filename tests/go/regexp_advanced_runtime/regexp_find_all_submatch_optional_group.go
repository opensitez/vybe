// vybe-test: go/regexp_advanced_runtime/regexp_find_all_submatch_optional_group
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

func main() { re := regexp.MustCompile(`colou?r`)
m := re.FindAllStringSubmatch("color colour", -1)
__check(fmt.Sprint(len(m)), "2")
__check(fmt.Sprint(m[1][0]), "colour") }
