// vybe-test: go/lang_interfaces_embedding/iface_eq_nil_both
// origin: languages/go/tests/go/test_lang_interfaces_embedding.rs

package main
import "fmt"
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { var i interface{}
__check(fmt.Sprint(i == nil), "true") }
