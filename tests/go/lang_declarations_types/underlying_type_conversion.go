// vybe-test: go/lang_declarations_types/underlying_type_conversion
// origin: languages/go/tests/go/test_lang_declarations_types.rs

package main
import "fmt"
type MyInt int
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { var m MyInt = 3
__check(fmt.Sprint(int(m)), "3") }
