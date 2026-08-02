// vybe-test: go/lang_declarations_types/nested_anonymous_struct
// origin: languages/go/tests/go/test_lang_declarations_types.rs

package main
import "fmt"
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { v := struct{ N int }{N: 8}
__check(fmt.Sprint(v.N), "8") }
