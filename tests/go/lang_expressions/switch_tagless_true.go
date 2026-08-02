// vybe-test: go/lang_expressions/switch_tagless_true
// origin: languages/go/tests/go/test_lang_expressions.rs

package main
import "fmt"
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { switch { case 1 < 2: __check(fmt.Sprint("t"), "t") } }
