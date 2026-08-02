// vybe-test: go/lang_declarations_types/string_index_mutation_illegal_use_byte_slice
// origin: languages/go/tests/go/test_lang_declarations_types.rs

package main
import "fmt"
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { b := []byte("ab")
b[0] = 'x'
__check(fmt.Sprint(string(b)), "xb") }
