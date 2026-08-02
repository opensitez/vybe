// vybe-test: go/lang_declarations_types/struct_field_exported
// origin: languages/go/tests/go/test_lang_declarations_types.rs

package main
import "fmt"
type P struct { X int }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { __check(fmt.Sprint(P{X:1}.X), "1") }
