// vybe-test: go/lang_declarations_types/int_to_string_conversion
// origin: languages/go/tests/go/test_lang_declarations_types.rs

package main
import "fmt"
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { s := fmt.Sprint(42)
__check(fmt.Sprint(s), "42") }
