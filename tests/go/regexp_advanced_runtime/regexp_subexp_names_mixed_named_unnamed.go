// vybe-test: go/regexp_advanced_runtime/regexp_subexp_names_mixed_named_unnamed
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

func main() { re := regexp.MustCompile(`(?P<id>\d+)-(\w+)`)
names := re.SubexpNames()
__check(fmt.Sprint(names[1]), "id")
__check(fmt.Sprint(names[2] == ""), "true") }
