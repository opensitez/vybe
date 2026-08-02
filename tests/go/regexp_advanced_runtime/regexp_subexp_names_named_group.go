// vybe-test: go/regexp_advanced_runtime/regexp_subexp_names_named_group
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

func main() { re := regexp.MustCompile(`(?P<year>\d{4})`)
names := re.SubexpNames()
__check(fmt.Sprint(names[1]), "year")
__check(fmt.Sprint(len(names)), "2") }
