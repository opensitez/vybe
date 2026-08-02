// vybe-test: go/strings_ops_extended/fields_func_splits_on_punctuation
// origin: languages/go/tests/go/test_strings_ops_extended.rs

package main
import "fmt"
import "strings"
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { f := strings.FieldsFunc("  a,b  c", func(r rune) bool { return r == ' ' || r == ',' })
__check(fmt.Sprint(len(f)), "3")
__check(fmt.Sprint(f[0]), "a")
__check(fmt.Sprint(f[2]), "c") }
