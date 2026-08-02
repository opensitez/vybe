// vybe-test: go/lang_interfaces_embedding/method_expression_pointer
// origin: languages/go/tests/go/test_lang_interfaces_embedding.rs

package main
import "fmt"
type N int
func (n *N) Inc() { *n++ }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { var v N = 1
f := (*N).Inc
f(&v)
__check(fmt.Sprint(v), "2") }
