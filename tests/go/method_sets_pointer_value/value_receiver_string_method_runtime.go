// vybe-test: go/method_sets_pointer_value/value_receiver_string_method_runtime
// origin: languages/go/tests/go/test_method_sets_pointer_value.rs

package main
import "fmt"
type label struct { text string }
func (l label) upper() string { return l.text + "!" }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { __check(fmt.Sprint(label{text: "hi"}.upper()), "hi!") }
