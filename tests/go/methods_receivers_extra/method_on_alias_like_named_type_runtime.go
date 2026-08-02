// vybe-test: go/methods_receivers_extra/method_on_alias_like_named_type_runtime
// origin: languages/go/tests/go/test_methods_receivers_extra.rs

package main
import "fmt"
type text string
func (t text) label() string { return string(t) + "!" }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { var value text = "go"
__check(fmt.Sprint(value.label()), "go!")
}
