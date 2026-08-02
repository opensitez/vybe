// vybe-test: go/lang_interfaces_embedding/method_expression_value
// origin: languages/go/tests/go/test_lang_interfaces_embedding.rs

package main
import "fmt"
type N int
func (n N) Inc() N { return n+1 }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { f := N.Inc
__check(fmt.Sprint(f(2)), "3") }
