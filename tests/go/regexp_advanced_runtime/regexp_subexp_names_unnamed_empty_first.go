// vybe-test: go/regexp_advanced_runtime/regexp_subexp_names_unnamed_empty_first
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
names := re.SubexpNames()
__check(fmt.Sprint(names[0] == ""), "true")
__check(fmt.Sprint(names[1] == ""), "true") }
