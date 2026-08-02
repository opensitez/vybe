// vybe-test: go/lang_declarations_types/var_block_declaration
// origin: languages/go/tests/go/test_lang_declarations_types.rs

package main
import "fmt"
var ( a = 1; b = 2 )
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { __check(fmt.Sprint(a + b), "3") }
