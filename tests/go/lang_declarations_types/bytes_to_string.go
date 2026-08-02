// vybe-test: go/lang_declarations_types/bytes_to_string
// origin: languages/go/tests/go/test_lang_declarations_types.rs

package main
import "fmt"
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { s := string([]byte{97})
__check(fmt.Sprint(s), "a") }
