// vybe-test: go/strings_ops_extended/fields_func_empty_string
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

func main() { f := strings.FieldsFunc("   ", func(r rune) bool { return r == ' ' })
__check(fmt.Sprint(len(f)), "0") }
