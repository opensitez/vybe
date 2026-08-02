// vybe-test: go/lang_builtins_control/struct_literal_field_order
// origin: languages/go/tests/go/test_lang_builtins_control.rs

package main
import "fmt"
type P struct { A int
B int }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { __check(fmt.Sprint(P{B:2, A:1}.A), "1") }
