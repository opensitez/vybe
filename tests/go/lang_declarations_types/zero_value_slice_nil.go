// vybe-test: go/lang_declarations_types/zero_value_slice_nil
// origin: languages/go/tests/go/test_lang_declarations_types.rs

package main
import "fmt"
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { var s []int
__check(fmt.Sprint(s == nil), "true") }
