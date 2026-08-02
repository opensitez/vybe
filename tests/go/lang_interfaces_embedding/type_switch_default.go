// vybe-test: go/lang_interfaces_embedding/type_switch_default
// origin: languages/go/tests/go/test_lang_interfaces_embedding.rs

package main
import "fmt"
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { switch any(1).(type) { default: __check(fmt.Sprint("d"), "d") } }
