// vybe-test: go/strings_builder/replacer_no_match_passthrough
// origin: languages/go/tests/go/test_strings_builder.rs

package main
import "fmt"
import "strings"
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { rep := strings.NewReplacer("z", "q")
__check(fmt.Sprint(rep.Replace("go")), "go") }
