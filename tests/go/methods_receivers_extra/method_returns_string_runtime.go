// vybe-test: go/methods_receivers_extra/method_returns_string_runtime
// origin: languages/go/tests/go/test_methods_receivers_extra.rs

package main
import "fmt"
type label struct { text string }
func (l label) value() string { return l.text }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { __check(fmt.Sprint(label{text: "vybe"}.value()), "vybe")
}
