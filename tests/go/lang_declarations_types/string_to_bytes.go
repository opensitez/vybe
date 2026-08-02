// vybe-test: go/lang_declarations_types/string_to_bytes
// origin: languages/go/tests/go/test_lang_declarations_types.rs

package main
import "fmt"
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { b := []byte("go")
__check(fmt.Sprint(len(b)), "2") }
