// vybe-test: go/strings_builder/replacer_ordered_pairs
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

func main() { rep := strings.NewReplacer("a", "b", "b", "c")
__check(fmt.Sprint(rep.Replace("ab")), "bc") }
