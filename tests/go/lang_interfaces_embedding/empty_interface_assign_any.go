// vybe-test: go/lang_interfaces_embedding/empty_interface_assign_any
// origin: languages/go/tests/go/test_lang_interfaces_embedding.rs

package main
import "fmt"
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { var a any = 5
__check(fmt.Sprint(a), "5") }
