// vybe-test: go/lang_expressions/struct_pointer_field_arrow
// origin: languages/go/tests/go/test_lang_expressions.rs

package main
import "fmt"
type P struct { N int }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { p := &P{N:1}
p.N = 2
__check(fmt.Sprint(p.N), "2") }
