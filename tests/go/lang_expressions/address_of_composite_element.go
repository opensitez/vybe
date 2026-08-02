// vybe-test: go/lang_expressions/address_of_composite_element
// origin: languages/go/tests/go/test_lang_expressions.rs

package main
import "fmt"
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { s := []int{1}
p := &s[0]
*p = 9
__check(fmt.Sprint(s[0]), "9") }
