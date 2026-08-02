// vybe-test: go/lang_declarations_types/package_level_init_expr
// origin: languages/go/tests/go/test_lang_declarations_types.rs

package main
import "fmt"
var n = len([]int{1,2,3})
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { __check(fmt.Sprint(n), "3") }
