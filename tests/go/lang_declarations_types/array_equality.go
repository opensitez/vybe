// vybe-test: go/lang_declarations_types/array_equality
// origin: languages/go/tests/go/test_lang_declarations_types.rs

package main
import "fmt"
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { __check(fmt.Sprint([2]int{1,2} == [2]int{1,2}), "true") }
