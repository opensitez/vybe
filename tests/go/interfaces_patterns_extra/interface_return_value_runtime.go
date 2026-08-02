// vybe-test: go/interfaces_patterns_extra/interface_return_value_runtime
// origin: languages/go/tests/go/test_interfaces_patterns_extra.rs

package main
import "fmt"
type speaker interface { speak() string }
type cat struct{}
func (cat) speak() string { return "meow" }
func build() speaker { return cat{} }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { __check(fmt.Sprint(build().speak()), "meow")
}
